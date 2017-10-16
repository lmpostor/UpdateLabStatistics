import gspread
from oauth2client.service_account import ServiceAccountCredentials

import imaplib
import email
import uu
import re
import csv
import os
import win32com.client
import win32api
import datetime

def sendstationstats(username,password,email):
	shell = win32com.client.Dispatch("WScript.Shell")
	path = "C:\\Optifacts.lnk" #hardcoded Optifacts shortcut.
	os.startfile(path)
	win32api.Sleep(1500)
	shell.SendKeys (<labname>) # example for the Dayton Lab I use "Day" as this will select it
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(2500)
	shell.SendKeys (username) #again this is hardcoded so that it is a regular username/logon
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(2500)
	shell.SendKeys (password)
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(5000)
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(1000)
	shell.SendKeys ("5")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("10")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("28")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("2")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("2")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(1000)
	shell.SendKeys ("1")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys (email)
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(1000)
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("0")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("0")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("0")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("0")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("11")
	shell.SendKeys ("{ENTER}")
	win32api.Sleep(500)
	shell.SendKeys ("13")
	shell.SendKeys ("{ENTER}")

def downloadfromEmail(address,password,label,endpath):
	imaplib._MAXLINE = 100000
	m = imaplib.IMAP4_SSL('imap.gmail.com')
	m.login(address,password)
	m.list()
	m.select(label) #you need to create a specifc label that optifacts reports need to be sorted under.

	result, items = m.search(None,'UNSEEN') #only pulls the latest email that is unread, in order to not use previous incomplete/outdated data
	items = items[0].split()


	if len(items) == 0:
		return False
	else:
		counter = 0

		for emailid in items:
			resp, data = m.fetch(emailid, "(RFC822)")
			body = data[0][1]
			mail = email.message_from_bytes(body).as_string()
			texts = re.search('begin(.*)end',mail,re.DOTALL)

			f = open("temp.txt","w")
			f.write(str(texts.group(0)))
			f.close()

			attachments = re.split('end',str(texts.group(0)),re.DOTALL) #manually splits into 3 different files based on a characteristic of UU encoding.


			for num in attachments:
				if counter == 2:
					path = "hourlystats.txt"
					fp = open(path, 'w')
					fp.write(num + "end")
					fp.close()

					uu.decode(path,endpath)
					os.remove(path)
				counter+=1
		os.remove("temp.txt")
		return True

if __name__ == "__main__":
	# use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)

	path1 = "hourlystats.csv"

	hourtest = datetime.datetime.now().hour


	if hourtest != 0:

		# Find a workbook by name and open the first sheet
		# Make sure you use the right name here.
		sheet = client.open(<worksheet_name>).get_worksheet(0)

		sendstationstats(<username>,<password>,<email>)

		win32api.Sleep(600000)

		if downloadfromEmail(<email>,<twofa_secret_key>,<email_sort_label>,path1):
			data = []

			with open(path1,newline='') as csvfile:
				datareader = csv.reader(csvfile,skipinitialspace=True,delimiter=',', quotechar='', quoting=csv.QUOTE_NONE)

				for row in datareader:
					if row:
						if int(float(row[1])) != hourtest:
							data.append(row)
			monthtodate = sheet.col_values(2)


			counter = 1
			newday = 0

			todaydate = data[0][0].split("/")
			formatdate = []


			for x in range(3):
				formatdate.append(todaydate[x].lstrip("0"))

			todaydate = "/".join(formatdate)
			print(todaydate)

			for datacell in monthtodate:
				if datacell == todaydate:
					break
				elif datacell == "":
					break
				counter+=1

			rowcount = 0
			columncount = 0


			datalist = sheet.range(counter,2,counter + len(data) - 1,6)


			for cell in datalist:
				cell.value = data[rowcount][columncount]
				columncount +=1
				if columncount == 5:
					columncount = 0
					rowcount+=1

			sheet.update_cells(datalist)

		else:
			win32api.Sleep(900000)
			if downloadfromEmail(<email>,<secret_key>,<email_sort_label>,path1):
				data = []

				with open(path1,newline='') as csvfile:
					datareader = csv.reader(csvfile,skipinitialspace=True,delimiter=',', quotechar='', quoting=csv.QUOTE_NONE)

					for row in datareader:
						if row:
							if int(float(row[1])) != hourtest:
								data.append(row)

				monthtodate = sheet.col_values(2)


				counter = 1
				newday = 0

				todaydate = data[0][0].split("/")
				formatdate = []


				for x in range(3):
					formatdate.append(todaydate[x].lstrip("0"))

				todaydate = "/".join(formatdate)

				for datacell in monthtodate:
					if datacell == todaydate:
						break
					elif datacell == "":
						break
					counter+=1

				rowcount = 0
				columncount = 0


				datalist = sheet.range(counter,2,counter + len(data) - 1,6)


				for cell in datalist:
					cell.value = data[rowcount][columncount]
					columncount +=1
					if columncount == 5:
						columncount = 0
						rowcount+=1

				sheet.update_cells(datalist)
