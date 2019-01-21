"""
Objective: Download attachments from unread emails within a specified outlook folder to the directory that this script is executing from
"""

import win32com.client
from win32com.client import Dispatch
import os.path
from pywintypes import com_error
import sys

def main():
	print("\nInitializing Script to Extract Attachments from Outlook emails...\n")
	print("To begin, first specify where your email(s) of interest are located")
	print(" - For default you can type Inbox")
	print(" - If the folder is nested, please include the full path e.g. To Do\\High Priority\n")

	folderPath = input("Please specify the location of the email(s): ")

	outlook = win32com.client.Dispatch("Outlook.Application")
	namespace = outlook.GetNamespace("MAPI")
	rootFolder = namespace.Folders.Item(1)
	inbox = getFolder(rootFolder, folderPath)

	for message in inbox.Items:
		if (message.Unread == True):
			print("Extracting attachments from", message.subject, message.CreationTime)
			getAttachments(message)
			message.Unread = False

	print("Done.")


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


def getAttachments(msg):
	"""Download attachments from a particular email"""

	for att in msg.Attachments:
		fileName = validateFileName(att.FileName)
		att.SaveASFile(fileName)


def validateFileName(fileName):
	"""Check if file exists. If true, add [#] to the beginning of the file name"""

	if os.path.isfile(fileName):
		count = 1
		while os.path.isfile("[" + str(count) + "]" + fileName):
			count = count + 1
		return (os.getcwd() + "\\" + "[" + str(count) + "]" + fileName)
	else:
		return (os.getcwd() + "\\" + fileName)


def printEmailsInFolder(folder):
	"""For troubleshooting"""

	print("Emails found in", folder, ": ")
	for message in folder.Items:
		print(message.subject, message.CreationTime)


main()