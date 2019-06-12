"""
Objective: Extract IOCs from unread Cofense emails within a specified outlook folder to the directory that this script is executing from
"""

import win32com.client
from win32com.client import Dispatch
import os.path
from pywintypes import com_error
import sys
import re

"""Temporary solution to simplify execution of the script"""
defaultFolderPath = ""	# Specify the location of the email in question
defaultSenderEmailAddress = ""	# Specify the email address of the sender to only look at messages from that address within the specified folder
defaultSubject = ""	# Specify the subject to only look at emails with that subject within the specified folder
""""""

def main():
	print("\nInitializing dcript to process Outlook emails...\n")

	folderPath = requestFolderPath(defaultFolderPath)

	outlook = win32com.client.Dispatch("Outlook.Application")
	namespace = outlook.GetNamespace("MAPI")
	rootFolder = namespace.Folders.Item(1)
	inbox = getFolder(rootFolder, folderPath)

	for message in inbox.Items:
		if (defaultSenderEmailAddress == "") | (message.SenderEmailAddress == defaultSenderEmailAddress):
			if (defaultSubject == "") | (message.Subject == defaultSubject):
				if (message.Unread == True):
					print("Extracting IOCs from", message.subject, message.CreationTime)
					parseEmail(message.body)
					message.Unread = False

	print("Done.")


def requestFolderPath(folderPath):
	"""Prompt user for folder path when defaultFolderPath is blank"""

	if folderPath == "":
		print("To begin, first specify where your email(s) of interest are located")
		print(" - For default you can type Inbox")
		print(" - If the folder is nested, please include the full path e.g. To Do\\High Priority\n")

		return input("Please specify the location of the email(s): ")
	else:
		return folderPath


def getFolder(baseFolder, folderPath):
	"""Parse user input and attempt to navigate to the specified folder"""

	temp = folderPath.split("\\", 1)
	try:
		if folderPath is "":
			return baseFolder
		else:
			if len(temp) == 1:
				return baseFolder.Folders[folderPath]
			else:
				return getFolder(baseFolder.Folders[temp[0]], temp[1])
	except com_error as e:
		if e.excepinfo[5] == -2147221233:
			print('The specified object could not be found. Are you sure you entered the correct folder path?')
		else:
			raise e
		sys.exit()


def parseEmail(msg):
	print("Commencing parsing of email...\n")

	body = cleanOriginalEmail(msg)

	rIOCs = [r'(Malicious File\(s\)\:)', r'(Malicious URL\:)', r'(Associated IP\:)']

	combinedR = re.compile('|'.join(rIOCs))

	bodySplitByIOCs = list(filter(None, re.split(combinedR, body)))

	index = 0

	for section in bodySplitByIOCs:
		if (index % 2 == 1):
			index += 1
			continue

		m0 = re.compile(rIOCs[0]).match(section)
		m1 = re.compile(rIOCs[1]).match(section)
		m2 = re.compile(rIOCs[2]).match(section)

		if m0:
			parseMaliciousFiles(bodySplitByIOCs[index+1])
		elif m1:
			parseMaliciousURLs(bodySplitByIOCs[index+1])
		elif m2:
			parseMaliciousIPs(bodySplitByIOCs[index+1])
		else:
			print("No IOCs found")

		index += 1

	print("\nParsing complete...\n")

def cleanOriginalEmail(msg):
	body = re.sub(r'\r', '', msg)
	body = re.sub(r'\n{2,}', '\n', body)

	pStart = re.compile(r'(Indicators of Compromise \(IOCs\)\:\n)')
	pEnd = re.compile(r'(\nCofense\nPhishing Defense Center\nphishing.defense@cofense.com)')

	try:
		body = pStart.split(body)[2]
	except:
		print("Declaration of IOCs not found.")
		exit()

	return pEnd.split(body)[0]

def sanitize(msg):
	cleanedMsg = re.sub(r'\[\.\]', ".", msg)

	return re.sub(r'\.', "[.]", cleanedMsg)

def parseMaliciousFiles(msg):
	rName = re.compile(r'(?<=(File Name\: )).*')
	rMD5 = re.compile(r'(?<=(MD5\: )).*')
	rSHA256 = re.compile(r'(?<=(SHA256\: )).*')

	rFile = [r'(?<=(File Name\: )).*', r'(?<=(MD5\: )).*', r'(?<=(SHA256\: )).*']

	combinedR = re.compile( '|'.join(rFile) )

	fileSplitByInfo = combinedR.finditer(msg)

	infoType = ["File Name", "MD5", "SHA256"]
	index = 0
	for info in fileSplitByInfo:
		if (index % 3 == 1):
			print(infoType[1] + ": " + info.group())
		elif (index % 3 == 2):
			print(infoType[2] + ": " + info.group())
		else:
			print(infoType[0] + ": " + info.group())
		index = index + 1

def parseMaliciousURLs(msg):
	rUrl = re.compile(r'(?=[a-zA-Z]{4,5}:\/\/).*')

	urls = rUrl.findall(msg)

	listOfDomains = []

	for url in urls:
		print("URL:", url)
		print("Domain:", extractDomain(url))

def extractDomain(url):
	mUrlProtocol = re.search(r'(((hxxp:\/\/)|(hxxps:\/\/)|(http:\/\/)|(https:\/\/))(((www)?((\[\.\])|\[\.\]))|\.?))', url).group(1)

	url = re.sub(re.escape(mUrlProtocol), "", url)

	try:
		mUrlPath = re.search(r'(?:((\[\.\]|\.)[a-z0-9]{1,5}))(\:[0-9]+)?(\/.*)', url).group(4)
		url = re.sub(re.escape(mUrlPath), "", url)
	except:
		pass

	return sanitize(url)

def parseMaliciousIPs(msg):
	msg = re.sub(r'\[\.\]', ".", msg)

	rIP = re.compile(r'[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}')

	IPs = rIP.findall(msg)

	for IP in IPs:
		print("IP:", sanitize(IP))

main()