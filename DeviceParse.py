#!/usr/bin/python
from ciscoconfparse import CiscoConfParse
# from ciscoconfparse.ccp_util import IPv4Obj
import re
import xlwt
from datetime import datetime
import os
import sys, getopt
import argparse


###############################################################
#                   Program Information
#
# Author:         Randy Afdahl
# Program name:   DeviceParse
# Purpose:        Parse switch and router files 
#                 creating an excel file with the output
#                 uses CiscoConfParse
#
# Revision History:
# Rev 1.0         Initial Release
#
###############################################################


proc = 0
print 

# parse the command line arguments
arg_parser = argparse.ArgumentParser()
arg_parser.add_argument('-t', required=True, action='store', dest='type_value', help='Type of Device ios, nxos or asa')
arg_parser.add_argument('-i', required=True, action='store', dest='input_value', help='Input Directory to Scan')
arg_parser.add_argument('-o', required=True, action='store', dest='output_value', help='Output xls file')
arg_parser.add_argument('--version','-v', action='version', version='%(prog)s 1.0')
results = arg_parser.parse_args()


#setup the styles for the excel worksheets
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
style3 = xlwt.easyxf('align: vert center, horiz center', 'font: name Times New Roman, color-index black')
style2 = xlwt.easyxf('font: name Times New roman;' 'align: wrap on, vert center, horiz left;' 'borders: left thin, right thin, top thin, bottom thin;' ,num_format_str='#,##0')
style4 = xlwt.easyxf('font: name Times New roman, bold True, height 250;' 'align: wrap on, vert center, horiz center;' 'pattern: pattern solid, fore_colour aqua;' 'borders: left thin, right thin, top thin, bottom thin;' ,num_format_str='#,##0')
style5 = xlwt.easyxf('font: name Times New roman;' 'align: vert center, horiz left;' 'borders: left thin, right thin, top thin, bottom thin;' 'pattern: pattern solid, fore_colour red;',num_format_str='#,##0')
wb = xlwt.Workbook()


# define the coloumn width in excel
col_width_type = 256 * 20         # 20 characters wide
col_width_port = 256 * 10         # 10 characters wide
col_width_status = 256 * 12       # 12 characters wide
col_width_vlan = 256 * 25         # 25 characters wide
col_width_ip = 256 * 15           # 15 characters wide
col_width_subnet = 256 * 15       # 15 characters wide
col_width_desc = 256 * 50         # 50 characters wide


# Here we grab all the files in the config directory
rootDir = results.input_value

for dirName, subdirList, fileList in os.walk(rootDir, topdown=False):
	for fname in fileList:
		DeviceFile = ('\t%s' % fname)             
		DeviceFile1 = ', '.join(re.findall(r'\S+$', DeviceFile))          
		
        

		# Parse the device configuration, need to use the proper type in ciscoconfparse for the OS on the switch
		# Here defined are the 3 types used in this script IOS, NXOS(nexus) and ASA 
		if results.type_value == ('asa'):
			parse = CiscoConfParse("%s/%s" % (rootDir, DeviceFile1), factory=True, syntax='ios')
		elif results.type_value == ('nxos'):
			parse = CiscoConfParse("%s/%s" % (rootDir, DeviceFile1), factory=True, syntax='nxos')
		elif results.type_value == ('ios'):
			parse = CiscoConfParse("%s/%s" % (rootDir, DeviceFile1), factory=True, syntax='ios')
		else:
            		print "You must use the type (-t) ios nxos or asa ONLY!"
			print "Exiting program"
			print ""
			exit()

		# Set the sheet name, have to grab the hostname
		host = parse.find_objects(r'hostname')[0]

		# Pull the hostname out of the list .... fun fun
		sheet_name = ', '.join(re.findall(r'\S+$', host.text))
		
		ws = wb.add_sheet(sheet_name.strip('\"'),cell_overwrite_ok=True)
                print "Processing ... ",sheet_name.strip('\"')
                proc = proc + 1		

		# set the widths from the definitions from above
		ws.col(0).width = col_width_type
		ws.col(1).width = col_width_port
		ws.col(2).width = col_width_status
		ws.col(3).width = col_width_vlan
		ws.col(4).width = col_width_ip
		ws.col(5).width = col_width_subnet
		ws.col(6).width = col_width_desc
		
		row=0
		
		ws.write(row, 0, "Hostname", style2)
		#ws.write_merge(row,row,1,2, re.findall(r'\S+$', host.text), style2)
		ws.write_merge(row,row,1,2, re.findall(r'\S+$', sheet_name.strip('\"')), style2)
                row= row+1
		
		version = parse.find_objects(r'ersion')[0]
		ws.write(row, 0, "Software Version", style2)
		ws.write_merge(row,row,1,2, re.findall(r'\S+\d', version.text), style2)
		row= row+1		
		
                try:
		     dns = parse.find_objects(r'domain-name')[0]
		     ws.write(row, 0, "Domain Name", style2)
		     ws.write_merge(row,row,1,2, re.findall(r'\S+$', dns.text), style2)
		     #print dns.text
		     #print re.findall(r'\S+$', dns.text)
		     row= row+1
                except:
                     errorspace = 1

		
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


                  
                # create the header row for VLAN INFO
		row = row+1
		ws.write(row, 0, "Vlan ID", style4)
		ws.write_merge(row,row,1, 3, "Name", style4)


                vlan_info = parse.find_objects(r'^vlan [0-9]+')
                for vlan_info_obj in vlan_info:
                     row = row+1
		     ws.write(row, 0, re.findall(r'[0-9]+', vlan_info_obj.text), style2)
                     for vlan_name in vlan_info_obj.re_search_children(r"^\s+name [aA-zZ]+"):
                          ws.write_merge(row,row,1,3, vlan_name.text, style2)
		     num = num+1


	
		# create the header row for INTERFACE INFO
		row = row+2
		ws.write(row, 0, "Interface Type", style4)
		ws.write(row, 1, "Port", style4)
		ws.write(row, 2, "Status", style4)
		ws.write(row, 3, "vlan", style4)
		ws.write(row, 4, "IP Address", style4)
		ws.write(row, 5, "Subnet Mask", style4)
		ws.write(row, 6, "Description", style4)


		interfaces = parse.find_objects('^interface')
		for interface_obj in interfaces:
		     row = row+1
		     if interface_obj.is_intf == 1:
		          if interface_obj.port_type == "port":
                               ws.write(row, 0 , "Ethernet", style2)
                               ws.write(row, 1, re.findall(r'[A-Z]+[0-9]+', interface_obj.text), style2)
                          else:
                               ws.write(row, 0, interface_obj.port_type, style2)
                               ws.write(row, 1, interface_obj.subinterface_number, style2)
		          ws.write(row, 4, interface_obj.ipv4_addr, style2)
		          ws.write(row, 5, interface_obj.ipv4_netmask, style2)
		          ws.write(row, 6, interface_obj.description, style2)
		          if interface_obj.access_vlan != 1:
		               ws.write(row, 3, str(interface_obj.access_vlan), style2)
		          elif interface_obj.access_vlan == 1: 
			       trunk_vlan = re.findall(r'[0-9]+,*', str(interface_obj.trunk_vlans_allowed))
		               ws.write(row, 3, trunk_vlan, style2)
			  elif interface_obj.re_search_children(r"^\s+shutdown"):
			       int_status = "shutdown"
                               ws.write(row, 2, int_status, style5)
                          else: 
                               ws.write(row, 2, "", style2)
                          
                               


################################################### FIREWALL ADDITIONS #########################################################
###################################################     UNDER TEST     #########################################################
		if results.type_value == 'asa':
			# create the header row for Object network
			row = row+2
			ws.write_merge(row,row,0,1, "Object Network", style4)
			ws.write_merge(row,row,2,4, "Element", style4)
			ws.write_merge(row,row,5,6, "Description", style4)
	

			object_network = parse.find_objects(r'^object network')
		        for object_network_obj in object_network:
		             row = row+1
			     ws.write_merge(row,row,0,1, re.findall(r'\S+$', object_network_obj.text), style2)
			     
		             for object_info in object_network_obj.re_search_children(r'\w+'):
		                  if object_info.text.find("description") == 1:
				  	ws.write_merge(row,row,5,6, object_info.text, style2)
				  else:
					ws.write_merge(row,row,2,4, object_info.text, style2)
					ws.write_merge(row,row,5,6, "", style2)
				  
			 


			# create the header row for Object Group NETWORK
			row = row+2
			ws.write_merge(row,row,0,1, "Object Group Network", style4)
			ws.write_merge(row,row,2,4, "Element", style4)
			ws.write_merge(row,row,5,6, "Description", style4)
	

			object_group = parse.find_objects(r'^object-group network')
		        for object_group_obj in object_group:
		             row = row+1
			     ws.write_merge(row,row,0,1, re.findall(r'\b(?!object|group|-|network)\w+.', object_group_obj.text), style2)
			     #print re.findall(r'\b(?!object|group|-|network)\w+.', object_group_obj.text)
		             for object_group_info in object_group_obj.re_search_children(r'\w+'):
		                  if object_group_info.text.find("description") == 1:
				  	ws.write_merge(row,row,5,6, object_group_info.text, style2)
				  else:
					ws.write_merge(row,row,2,4, object_group_info.text, style2)
					row = row+1
					ws.write_merge(row,row,5,6, "", style2)
			     
			# create the header row for Object Group SERVICE
			row = row+2
			ws.write_merge(row,row,0,1, "Object Group Service", style4)
			ws.write_merge(row,row,2,4, "Element", style4)
			ws.write_merge(row,row,5,6, "Description", style4)
	

			object_group = parse.find_objects(r'^object-group service')
		        for object_group_obj in object_group:
		             row = row+1
			     ws.write_merge(row,row,0,1, re.findall(r'\b(?!object|group|-|service)\w+.', object_group_obj.text), style2)
			     #print re.findall(r'\b(?!object|group|-|service)\w+.', object_group_obj.text)
		             for object_group_info in object_group_obj.re_search_children(r'\w+'):
		                  if object_group_info.text.find("description") == 1:
				  	ws.write_merge(row,row,5,6, object_group_info.text, style2)
				  else:
					ws.write_merge(row,row,2,4, object_group_info.text, style2)
					row = row+1
					ws.write_merge(row,row,5,6, "", style2)
		                    

			# create the header row for Access-List and Rules
			row = row+2
			ws.write_merge(row,row,0,1, "Access list", style4)
			ws.write_merge(row,row,2,4, "Element", style4)
			ws.write_merge(row,row,5,6, "Description", style4)
	

			object_access_list = parse.find_objects(r'^access-list')
		        for object_access_list_obj in object_access_list:
			     # get access list
			     access_list = re.findall(r'\b(?!access|list|-)\S+', object_access_list_obj.text)
			     #print access_list[4]
		  	
			object_access_group = parse.find_objects(r'^access-group')
		        for object_access_group_obj in object_access_group:
			     # get access groups
			     access_group = re.findall(r'\b(?!access|group|-)\S+', object_access_group_obj.text)
			     print access_group
			 
################################################ END OF FIREWALL ADDITIONS ######################################################


wb.save(results.output_value)

# output of each devices process to give feedback to the user
print "----------------------------"
print "Total sheets processed = ", proc
print "Output File: ",results.output_value
print "----------------------------"
