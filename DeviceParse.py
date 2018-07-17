#!/usr/bin/python
from ciscoconfparse import CiscoConfParse
import re
import xlwt
from datetime import datetime
import os
import sys, getopt


#setup the excel worksheet
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
style3 = xlwt.easyxf('align: vert center, horiz center', 'font: name Times New Roman, color-index black')
style2 = xlwt.easyxf('font: name Times New roman;' 'align: vert center, horiz left;' 'borders: left thin, right thin, top thin, bottom thin;' ,num_format_str='#,##0')
style4 = xlwt.easyxf('font: name Times New roman, bold True, height 250;' 'align: vert center, horiz center;' 'pattern: pattern solid, fore_colour aqua;' 'borders: left thin, right thin, top thin, bottom thin;' ,num_format_str='#,##0')
style5 = xlwt.easyxf('font: name Times New roman;' 'align: vert center, horiz left;' 'borders: left thin, right thin, top thin, bottom thin;' 'pattern: pattern solid, fore_colour red;',num_format_str='#,##0')
wb = xlwt.Workbook()


# define the coloumn width
col_width_type = 256 * 20         # 20 characters wide
col_width_port = 256 * 10         # 10 characters wide
col_width_status = 256 * 12       # 12 characters wide
col_width_vlan = 256 * 25         # 25 characters wide
col_width_ip = 256 * 15           # 15 characters wide
col_width_subnet = 256 * 15       # 15 characters wide
col_width_desc = 256 * 50         # 50 characters wide


# Here we grab all the files in the config directory
rootDir = 'config'
for dirName, subdirList, fileList in os.walk(rootDir, topdown=False):
	for fname in fileList:
		DeviceFile = ('\t%s' % fname)
		DeviceFile1 = ', '.join(re.findall(r'\S+$', DeviceFile))
		print DeviceFile1


		# Set the sheet name, have to grab the hostname
		parse= CiscoConfParse("config/%s" % DeviceFile1, factory=True)
		host = parse.find_objects(r'hostname')[0]
		# Pull the hostname out of the list .... fun fun
		sheet_name = ', '.join(re.findall(r'\S+$', host.text))
		ws = wb.add_sheet(sheet_name,cell_overwrite_ok=True)
                print sheet_name

		# set the widths
		ws.col(0).width = col_width_type
		ws.col(1).width = col_width_port
		ws.col(2).width = col_width_status
		ws.col(3).width = col_width_vlan
		ws.col(4).width = col_width_ip
		ws.col(5).width = col_width_subnet
		ws.col(6).width = col_width_desc

		parse= CiscoConfParse("config/%s" % DeviceFile1, factory=True)
		row=0

		host= parse.find_objects(r'hostname')[0]
		ws.write(row, 0, "Hostname", style2)
		ws.write_merge(row,row,1,2, re.findall(r'\S+$', host.text), style2)
		row= row+1

		version = parse.find_objects(r'version')[0]
		ws.write(row, 0, "Software Version", style2)
		ws.write_merge(row,row,1,2, re.findall(r'\S+$', version.text), style2)
		row= row+1


		dns = parse.find_objects(r'ip domain-name')[0]
		ws.write(row, 0, "Domain Name", style2)
		ws.write_merge(row,row,1,2, re.findall(r'\S+$', dns.text), style2)
		row= row+1

		ip_name_server = parse.find_objects(r'ip name-server')
		num = 1
		for ip_name_obj in ip_name_server:
		     ws.write(row, 0, ("Name servers (%d)" % num), style2)
		     ws.write_merge(row,row,1,2, re.findall(r'[0-9]+.[0-9]+.[0-9]+.[0-9]+', ip_name_obj.text), style2)
		     row= row+1
		     num = num+1

		logging_host = parse.find_objects(r'logging host')
		num = 1
		for logging_obj in logging_host:
		     ws.write(row, 0, ("Logging servers (%d)" % num), style2)
		     ws.write_merge(row,row,1,2, re.findall(r'[0-9]+.[0-9]+.[0-9]+.[0-9]+', logging_obj.text), style2)
		     row= row+1
		     num = num+1

		snmp_server = parse.find_objects(r'snmp-server host')
		num = 1
		for snmp_obj in snmp_server:
		     ws.write(row, 0, ("SNMP servers (%d)" % num), style2)
		     ws.write_merge(row,row,1,2, re.findall(r'[0-9]+.[0-9]+.[0-9]+.[0-9]+', snmp_obj.text), style2)
		     row= row+1
		     num = num+1

		ntp_server = parse.find_objects(r'ntp server')
		num = 1
		for ntp_obj in ntp_server:
		     ws.write(row, 0, ("NTP servers (%d)" % num), style2)
		     ws.write_merge(row,row,1,2, re.findall(r'[0-9]+.[0-9]+.[0-9]+.[0-9]+', ntp_obj.text), style2)
		     row= row+1
		     num = num+1

                vlan_info = parse.find_objects(r'^vlan [0-9]+')
                for vlan_info_obj in vlan_info:
        	     ws.write(row, 0, "VLAN IDs ", style2)
		     ws.write_merge(row,row,1,2, re.findall(r'[0-9]+', vlan_info_obj.text), style2)
                     if vlan_info_obj.re_search_children(r"^.name [aA-zZ]+"):
                          ws.write(row, 3, "test", style2)
                          #ws.write(row, 3, re.findall(r'^.name [aA-zZ]+', vlan_info_obj.text), style2)
		     row= row+1
		     num = num+1



		# create the header row
		row = row+1                                   #create a space
		ws.write(row, 0, "Interface Type", style4)
		ws.write(row, 1, "Port", style4)
		ws.write(row, 2, "Status", style4)
		ws.write(row, 3, "vlan", style4)
		ws.write(row, 4, "IP Address", style4)
		ws.write(row, 5, "Subnet Mask", style4)
		ws.write(row, 6, "Description", style4)


		interfaces= parse.find_objects('^interface')
		for interface_obj in interfaces:
		     row= row+1
		     if interface_obj.is_intf == 1:
		          ws.write(row, 0, interface_obj.port_type, style2)
		          ws.write(row, 1, interface_obj.subinterface_number, style2)
		          ws.write(row, 4, interface_obj.ipv4_addr, style2)
		          ws.write(row, 5, interface_obj.ipv4_netmask, style2)
		          ws.write(row, 6, interface_obj.description, style2)
		          if interface_obj.access_vlan != 1:
		               ws.write(row, 3, str(interface_obj.access_vlan), style2)
		          if interface_obj.access_vlan == 1:
			           trunk_vlan = re.findall(r'[0-9]+,*', str(interface_obj.trunk_vlans_allowed))
		               ws.write(row, 3, trunk_vlan, style2)
			  if interface_obj.re_search_children(r"^\s+shutdown"):
			       int_status = "shutdown"
                               ws.write(row, 2, int_status, style5)
                          else:
                               ws.write(row, 2, "", style2)
                          if interface_obj.re_search_children(r"^\s+ip address"):
			       print interface_obj.re_search_children


wb.save("example.xls")
