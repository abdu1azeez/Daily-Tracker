#!/usr/bin/python

#**************************************************************************************************************************************************************
# Script Version: Version 5.0
# Owner: M Abdul Azeez Muqthar (06121Z)
# Program name: 2_Send_Shift_Tracker
# Release Date: 11-Apr-2019
# Credits: Source languages used - Python, Html, Bash
# Version 5.0 updates: Process has been simplified ,html output is generated and report file is attached.
#
# Input:  From Service Now files and Manual updates
# Output: Mail Report, Report.xlsx and Email
#
#**************************************************************************************************************************************************************
#
#
#************************************************************* Please refrain from editing this file **********************************************************

import pandas as pd
import numpy as np
import glob
import subprocess
import os
import sys
from datetime import date, timedelta, datetime


to_address = "mw_imi@wwpdl.vnet.ibm.com, its_managed_services_middleware_support@wwpdl.vnet.ibm.com"		# To Address
cc_address = "elijosep@in.ibm.com, vp.shetty@in.ibm.com"							# CC Address	

h = datetime.now().hour
d = datetime.today().strftime('%d')

shift = 'Morning Shift' if(11 < h < 20) else 'Night Shift' if (2 < h < 12) else 'Noon Shift'

date_format = date.today().strftime('%a %b %dst %Y') if (d in ['1', '21', '31']) else date.today().strftime('%a %b %dnd %Y') if (d in ['2', '22']) else date.today().strftime('%a %b %drd %Y') if (d in ['3', '23']) else date.today().strftime('%a %b %dth %Y')

subject = shift + ' Tracker Report as on >>> ' + date_format


incident_columns = sorted([u'Number', u'Priority', u'Affected Date', u'Company', u'State', u'Resolved', u'Category', u'Assignment group', u'Assigned to', u'Resolved by', u'Work notes', u'Short description', u'Breached Resolution Time', u'Type'])

columns = ['Number' , 'Short description','Priority', 'Category', 'Affected Date', 'Company', 'State', 'Assigned to', 'Resolved by', 'Work notes']

sr_columns = columns[:]
sr_columns.remove('Category')

change_columns = sorted([u'Number', u'Company', u'Short description', u'Approval', u'Type', u'State', u'Planned start date', u'Planned end date', u'Assigned to'])

incident_files_count = int(subprocess.check_output('ls ./Input/|grep incident*|wc -l', shell=True).strip('\n'))
change_files_count = int(subprocess.check_output('ls ./Input/|grep change*|wc -l', shell=True).strip('\n'))
manual_update_file = subprocess.check_output('[ -f ./Input/Manual_Updates.xlsx ] && echo True || echo False', shell=True).strip('\n')


os.system('cd ./Input/; c=0;for f in incident*.xlsx; do ((c++));  mv "$f" incident\("$c"\).xlsx ; done > /dev/null 2>&1')	
os.system('cd ./Input/; c=0;for f in change*.xlsx; do ((c++));  mv "$f" change\("$c"\).xlsx ; done > /dev/null 2>&1')

duplicate_incidents = os.popen('cd Input; [ `diff -q incident\(1\).xlsx incident\(2\).xlsx|wc -l` = 0 ] || [ `diff -q incident\(2\).xlsx incident\(3\).xlsx|wc -l` = 0 ] || [ `diff -q incident\(1\).xlsx incident\(3\).xlsx|wc -l` = 0 ] && echo True || echo False').read().strip('\n')
duplicate_changes = os.popen('cd Input; [ `diff -q change\(1\).xlsx change\(2\).xlsx|wc -l` = 0 ] || [ `diff -q change\(2\).xlsx change\(3\).xlsx|wc -l` = 0 ] || [ `diff -q change\(1\).xlsx change\(3\).xlsx|wc -l` = 0 ] && echo True || echo False').read().strip('\n')

if (incident_files_count != 3):
	print('Terminated: Download incident files of all three regions')

elif (change_files_count != 3):
	print('Terminated: Download change files of all three regions')

elif (manual_update_file == 'False'):
	print('Terminated: Individual tracker file not found')

elif (duplicate_incidents == 'True'):
	print('Terminated: Duplicate Incidents found')

elif (duplicate_changes == 'True'):
	print('Terminated: Duplicate Changes found')

else:

	incident_data = pd.DataFrame()
	for f, region in zip(glob.glob('./Input/incident*'), ['AP', 'EU', 'NA']):
	   df = pd.read_excel(f)
	   if (sorted(df.columns) != incident_columns):
		sys.exit('Fields ' + str(list(set(incident_columns).difference(df.columns))[:]) + ' not found in ' + region + ' region')
	   incident_data = incident_data.append(df, ignore_index=True)
	
	o = incident_data.select_dtypes(include=['object']).columns.tolist()
	incident_data[o] = incident_data[o].fillna('')

	for col in o:
	    incident_data[col] = incident_data[col].apply(lambda x: x.encode('ascii','ignore'))

	incident_writer = pd.ExcelWriter('./Input/Incidents.xlsx', engine='xlsxwriter')
	incident_data.to_excel(incident_writer, sheet_name='Sheet1', index = False)
	incident_writer.save()
	os.system('rm -rf ./Input/incident*')


	change_data = pd.DataFrame()
	for f, region in zip(glob.glob('./Input/change*'), ['AP', 'EU', 'NA']):
	   df = pd.read_excel(f)
	   if (sorted(df.columns) != change_columns):
		sys.exit('Fields ' + str(list(set(change_columns).difference(df.columns))[:]) + ' not found in ' + region + ' region')
	   change_data = change_data.append(df, ignore_index=True)

	o = change_data.select_dtypes(include=['object']).columns.tolist()
	change_data[o] = change_data[o].fillna('')

	for col in o:
	    change_data[col] = change_data[col].apply(lambda x: x.encode('ascii','ignore'))

	change_writer = pd.ExcelWriter('./Input/Changes.xlsx', engine='xlsxwriter')
	change_data.sort_values(by = 'Planned start date', inplace= True)
	change_data.to_excel(change_writer, sheet_name='Sheet1', index = False)
	change_writer.save()
	os.system('rm -rf ./Input/change*')

	with pd.option_context('display.max_colwidth', -1) : 
		completed_changes = change_data[change_data['State'].isin(['Review', 'Closed'])].to_html(index = False, justify = 'center')
	
	with pd.option_context('display.max_colwidth', -1) : 
		upcoming_changes = change_data[change_data['State'].isin(['Review', 'Closed']) == False].to_html(index = False, justify = 'center')	
	
	incident_data['Work notes'].fillna('', inplace=True)
	incident_data['Work notes'] = incident_data['Work notes'].apply(lambda x: x.split('\n\n')[0])

	incident_data['Work notes'] = incident_data['Work notes'].apply(lambda x: x.replace('\n', '<br/>').replace('(Work notes)', '<br/>'))
	
	incident_data['Affected Only Date'] = incident_data['Affected Date'].apply(lambda x: x.date())


	Individual_Tracker_data = pd.read_excel('./Input/Manual_Updates.xlsx', sheet_name='Individual Tracker')
	Individual_Tracker_data['Hours'].fillna(9, inplace = True)
	Individual_Tracker_data['Account'].fillna('General', inplace = True)

	o = Individual_Tracker_data.select_dtypes(include=['object']).columns.tolist()
	Individual_Tracker_data[o] = Individual_Tracker_data[o].fillna('')

	for col in o:
	    Individual_Tracker_data[col] = Individual_Tracker_data[col].apply(lambda x: x.encode('ascii','ignore'))


	Individual_Tracker_data['Description'] = Individual_Tracker_data['Description'].apply(lambda x: x.replace('\n', '<br/>'))

	with pd.option_context('display.max_colwidth', -1):
		Individual_Tracker_data_table = Individual_Tracker_data.to_html(index=False, justify='center').replace('&lt;', '<').replace('&gt;', '>')


	Handover_data = pd.read_excel('./Input/Manual_Updates.xlsx', sheet_name='Handover')
	Handover_data.fillna('', inplace=True)


	o = Handover_data.select_dtypes(include=['object']).columns.tolist()
	Handover_data[o] = Handover_data[o].fillna('')

	for col in o:
	    Handover_data[col] = Handover_data[col].apply(lambda x: x.encode('ascii','ignore'))

	Handover_data_html = "\n<p><b>\n<font color='red'>\n"+\
	Handover_data['High Importance'].str.cat(sep = '<br/>').replace('\n', '<br/>')+\
	"\n</font>\n</p></b>\n<p>\n<font color='blue'>\n"+\
	Handover_data['Information'].str.cat(sep = '<br/>').replace('\n', '<br/>') +\
	"\n</font>\n</p>\n<p>"+\
	Handover_data['Low Importance'].str.cat(sep = '<br/>').replace('\n', '<br/>') + "\n</p>\n <hr>"

	

	with pd.option_context('display.max_colwidth', -1) : 
		SLA_Hold = incident_data[(incident_data['State'] == 'SLA Hold')& (incident_data['Type'] == 'Incident')][columns].sort_values(by = 'Affected Date', ascending = False).to_html(index = False, justify = 'center').replace('&lt;', '<').replace('&gt;', '>')

	with pd.option_context('display.max_colwidth', -1) : 
		In_Progress = incident_data[(incident_data['State'] == 'In Progress')& (incident_data['Type'] == 'Incident')][columns].sort_values(by = 'Affected Date', ascending = False).to_html(index = False, justify = 'center').replace('&lt;', '<').replace('&gt;', '>')

	with pd.option_context('display.max_colwidth', -1) : 
		Resolution_Breach = incident_data[incident_data['Resolved'] > incident_data['Breached Resolution Time']][columns].sort_values(by = 'Affected Date', ascending = False).to_html(index = False, justify = 'center').replace('&lt;', '<').replace('&gt;', '>')
	

	f = open('./Output/Mail_Report.html', 'w+')

	f.write("<html>\n<head>\n<title>Tracker Report</title>\n<style>\ntable, th, td {\n    border: 1px solid black;\n    border-collapse: collapse;\n }\n table {width:100%} \n th, td {\n    padding: 5px;\n}\n</style>\n</head>\n<body>")
	
	f.close()


	f = open('./Output/Mail_Report.html', 'a+')

	f.write('<p>Hi Team,</p>\n')

	f.write('<br/>\n<h4>Shift Handover / Pipeline Activities/ Major Updates:</h4>\n<hr>')
	
	f.write(Handover_data_html)

	f.write('<br/>\n<h4>Changes : </h4><br/>\n Completed :\n<br/>')

	f.write(completed_changes)
	
	f.write('\n<br/> Upcoming:\n<br/>')

	f.write(upcoming_changes)

	f.write('<br/>\n<br/><h4>SLA Hold :</h4>\n')
	
	f.write(SLA_Hold)

	f.write('<br/>\n<br/><h4>In Progress :</h4>\n')

	f.write(In_Progress)

	f.write('<br/>\n<br/><h4>Resolution Breach :</h4>\n')

	f.write(Resolution_Breach)

	f.write('<br/>\n<h4>Individual Tracker:</h4>\n')
	
	f.write(Individual_Tracker_data_table)	
	
	f.write('\n<br/><br/>\n*** This is an automatically generated email, please do not reply ***')

	f.write('</body></html>')

	f.close()


	report_writer = pd.ExcelWriter('./Output/Report.xlsx', engine='xlsxwriter')

	incident_data['Work notes'] = incident_data['Work notes'].apply(lambda x: x.replace('<br/>', ' '))

	incident_data[(incident_data['Affected Only Date'] == pd.Timestamp(date.today() - timedelta(1)).date()) | (incident_data['Affected Only Date'] == pd.Timestamp(date.today() - timedelta(0)).date()) & (incident_data['Type'] == 'Incident')][columns].sort_values(by = 'Affected Date', ascending = False).to_excel(report_writer, sheet_name= 'Incoming Incidents', index = False)

	incident_data[(incident_data['State'] == 'SLA Hold')& (incident_data['Type'] == 'Incident')][columns].sort_values(by = 'Affected Date', ascending = False).to_excel(report_writer, sheet_name= 'SLA Hold', index = False)

	incident_data[(incident_data['State'] == 'In Progress')& (incident_data['Type'] == 'Incident')][columns].sort_values(by = 'Affected Date', ascending = False).to_excel(report_writer, sheet_name= 'In Progress', index = False)

	incident_data[incident_data['Resolved'] > incident_data['Breached Resolution Time']][columns].sort_values(by = 'Affected Date', ascending = False).to_excel(report_writer, sheet_name= 'Resolution Breaches', index = False)

	incident_data[ (incident_data['Type'] == 'Service Request') & (incident_data['State'] != 'Resolved')][sr_columns].sort_values(by = 'Affected Date', ascending = False).to_excel(report_writer, sheet_name= 'Service Requests', index = False)	

	change_data.to_excel(report_writer, sheet_name= 'Changes', index = False)

	change_data[change_data['State'].isin(['Review', 'Closed'])].to_excel(report_writer, sheet_name= 'Completed Changes', index = False)

	change_data[change_data['State'].isin(['Review', 'Closed'])  == False].to_excel(report_writer, sheet_name= 'Upcoming Changes', index = False)


	report_writer.save()

	os.system("if [ `ps aux|grep postfix|wc -l` -lt 4 ]; then sudo start postfix; fi")
	os.system("gio open ./Output/Mail_Report.html")
	os.system("""echo 'Do you want to mail it? [Y/N]'; read answer;  if [ "$answer" != "${answer#[Yy]}" ] ;then cat ./Output/Mail_Report.html |/usr/local/bin/mutt -e 'set content_type=text/html' -s '""" + subject +"' -a ./Output/Report.xlsx -c '"+ cc_address + "' -- " + to_address + ";echo 'Mail Sent'; else echo 'Mail not sent'; fi")
	os.system("rm -rf ./Input/Incidents.xlsx ./Input/Changes.xlsx; sleep 3")
	

