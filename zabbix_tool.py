#coding:utf-8
# encoding=utf8  
from pyzabbix import ZabbixAPI
import xlrd
ZABBIX_SERVER = 'http://10.255.12.26/zabbix/'
zapi = ZabbixAPI(ZABBIX_SERVER)
zapi.login('admin', 'teatime')
import sys

#host_name = '10.247.100.4'

def createItem(hostid,keyname,keydesc,flag):
		#check if item for incoming traffic item already existed
	if flag == 'in':
		items= zapi.item.get(
			hostids = hostid,
			search={'key_': "["+keyname+"]", 'snmp_oid' : 'ifHCInOctets' }
		)
	if flag == 'out':
		items= zapi.item.get(
			hostids = hostid,
			search={'key_': "["+keyname+"]", 'snmp_oid' : 'ifHCOutOctets' }
		)	######print (items_in)
	if items:
		itemkey = items[0]["key_"] 
		print("Item {0} already exist".format(itemkey))
		itemid = items[0]["itemid"]
		return(itemid)
	else:
		if flag == 'in':
			itemkey = 'ifHCInOctets['+ keyname + ']'
			items = zapi.item.create(
				type=4,
				snmp_community='{$SNMP_COMMUNITY}',
				interfaceid=int_id,
				name='Incoming traffic on interface $1',
				key_=itemkey,
				hostid=host_id,
				snmp_oid='IF-MIB::ifHCInOctets["index","ifDescr","'+keyname+'"]',
				delay=60,
				value_type=3,
				units='bps',
				multiplier=1,
				delta=1,
				formula=8,
				description= keydesc
			)
		elif flag == 'out':
			itemkey = 'ifHCOutOctets['+ keyname + ']'
			items = zapi.item.create(
				type=4,
				snmp_community='{$SNMP_COMMUNITY}',
				interfaceid=int_id,
				name='Outgoing traffic on interface $1',
				key_=itemkey,
				hostid=host_id,
				snmp_oid='IF-MIB::ifHCOutOctets["index","ifDescr","'+keyname+'"]',
				delay=60,
				value_type=3,
				units='bps',
				multiplier=1,
				delta=1,
				formula=8,
				description= keydesc
			)
		print(items)
		itemid = items["itemids"][0]
	return itemid

def createTrigger(hostid,itemid,hostname,keyname,keydesc,threshold,flag):	
	triggers_list= zapi.trigger.get(
		hostids = hostid,
		itemids = itemid,
	) 
	if triggers_list:
		if flag == 'in':
			print("Incoming traffic trigger for interface {0} already exist".format(keyname)) 
		elif flag == 'out':
			print("Outgoing traffic trigger for interface {0} already exist".format(keyname))  
		
	else:
		if flag == 'in':					
			trigger_desc = "{HOST.NAME} : Incoming traffic on "+keyname+"("+keydesc+")"+\
							"exceed "+str(float(threshold)*100)+"% for the last 5 minutes!"
			itemkey = 'ifHCInOctets['+ keyname + ']'
		elif flag == 'out':
			trigger_desc = "{HOST.NAME} : Outgoing traffic on "+keyname+"("+keydesc+")"+\
				" exceed "+str(float(threshold)*100)+"% for the last 5 minutes!"
			itemkey = 'ifHCOutOctets['+ keyname + ']'

		trigger_expr = "("+"{"+hostname+":"+itemkey+"."+"delta(5m)}"+"/"+\
							"{"+hostname+":"+itemkey+"."+"avg(5m)}>"+str(threshold)+")"
		trigger_prio = 4
		triggers = zapi.trigger.create(
			description = trigger_desc,
			expression  = trigger_expr,
			priority = trigger_prio,
		)
		print("trigger created")
		print(triggers) 
		
def createGraph(groupname,hostid,iteminid,itemoutid,hostdesc,keyname,keydesc):
	graph_name = groupname +' - '+hostdesc+' - '+keyname+' - '+keydesc
	print(graph_name)
	graphs_list= zapi.graph.get(
		hostids = hostid,
		search={'name': graph_name }
	)
	print(graphs_list)
	if graphs_list:
		graph_id = graphs_list[0]["graphid"]
		print("here is         "+graph_id)
		graphs = zapi.graph.update(
			graphid = graph_id,
			width = 900,
			gitems = [{'itemid':itemoutid,'color':'002A97','drawtype':'0'},\
							{'itemid':iteminid,'color':'00CF00','drawtype':'1'}]
		)
		return graph_id
	else:
		graphs = zapi.graph.create(
			name = graph_name,
			width = 900,
			gitems = [{'itemid':itemoutid,'color':'002A97','drawtype':'0'},\
						{'itemid':iteminid,'color':'00CF00','drawtype':'1'}]
		)
		graph_id = graphs["graphids"][0]
		print(graph_id)
		return graph_id

def createScreen(screenname,graphid):

	screens = zapi.screen.get(
		output='extend',
		selectScreenItems='extend',
		filter = {"name": screenname},
	)
	if screens:
		screen_id = screens[0]["screenid"]
		screen_vsize = screens[0]["vsize"]
	else:
		screenids= zapi.screen.create(
		name = screenname,
		hsize = 1,
		vsize = 5,
		)
		screen_id = screenids["screenids"][0]
		screen_vsize = 5
	#get the number of screen item
	screenitems_count= zapi.screenitem.get(
		screenids = screen_id,
		sortorder = 'DESC',
   		output = 'count',
		countOutput = 'count'
	)

	if screen_vsize < screenitems_count :
		screen_vsize = str(int(screen_vsize) + 5)
		zapi.screen.update(
			screenid = screen_id,
			vsize = screen_vsize
		)
	#check if the graph already exist in the screen
	screenitems = zapi.screenitem.get(
		#output='extend',
		screenids = screen_id,
		filter={"resourceid":graphid},
		)
	#insert the graph if not exist
	if screenitems:
		print("Graph {0} already exist in screen".format(graph_id))
	else:
		screenitem_id= zapi.screenitem.create(
			screenid = screen_id,
			resourcetype = 0,
			resourceid = graphid,
			height = 200,
			width = 900,
			x = 0,
			y = screenitems_count
		)
		return screenitem_id

print(sys.argv[1])

#获取excel里的数据		
filename = sys.argv[1]
print(filename)
workbook = xlrd.open_workbook(filename)
table = workbook.sheets()[0]
nrows = table.nrows

host_desc = ""
key_name= ""
key_desc= ""
host_id = "1"
for row in range(nrows):
	print(row)
	if(row == 0):
		continue	
	host_name = table.cell(row,0).value
	hosts = zapi.host.get(
		selectGroups='extend', 
		selectInterfaces='extend',
		filter={"host": host_name}
	)
	if hosts:
		print(hosts)
		host_id = hosts[0]["hostid"]
		host_desc = hosts[0]["name"]
		group_name = hosts[0]["groups"][0]["name"]	
		int_id = hosts[0]["interfaces"][0]["interfaceid"]	
		######print (group_name)
		print("Found host id {0}".format(host_id))
	else:
		print("No hosts found")
		break
	#Get Key name (Interface name)
	key_name = table.cell(row,1).value
	print(host_id)
	print (key_name)
	#Get Key description (Interface description)
	key_desc = table.cell(row,2).value
	print (key_desc)
	T_in = table.cell(row,3).value
	print (str(T_in))
	T_out = table.cell(row,4).value
	print(str(T_out))
	
	item_in_id = createItem(host_id,key_name,key_desc,'in')	
	item_out_id = createItem(host_id,key_name,key_desc,'out')	
	graph_id = createGraph(group_name,host_id,item_in_id,item_out_id,host_desc,key_name,key_desc)
	if(graph_id != 0):
		screen_id = createScreen(host_desc,graph_id)
	if(T_in):	
		createTrigger(host_id,item_in_id,host_name,key_name,key_desc,T_in,'in')
	if(T_out):	
		createTrigger(host_id,item_out_id,host_name,key_name,key_desc,T_out,'out')



		
