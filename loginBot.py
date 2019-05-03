from slackclient import SlackClient
import pprint, openpyxl, os, excel2img, sys, time
from colorama import Fore, Back, Style, init


init(autoreset = True)

def sendMessage(message, realName, channel):
	'''
	input: text to send, name of user, channel
	output: text to be sent to channel user
	'''
	slack.api_call(
	"chat.postMessage",
	channel=channel,
	text= message)

def sendMistakeReport(name, channel):
	'''
	input: sheet, channel
	output: sends mistakeReportWorkbook[sheet] to user's channel
	'''
	# TODO - remove hardcoding of fileContent
	#fileContent = {'file': (f'Sent/{name}MistakeReport.xlsx', open(f'Sent/{name}MistakeReport.xlsx', 'rb'), 'xlsx')}
	#slack.api_call(
	#'files.upload',
	#channels = channel,
	#file = fileContent['file'],
	#title = f'{name} Mistake Report'
	#)
	# send png. Some users may not have the ability to open xlsx files (android, login terminals)
	fileContent = {'file': (f'Sent/{name}MistakeReport.png', open(f'Sent/{name}MistakeReport.png', 'rb'), 'png')}
	slack.api_call(
	'files.upload',
	channels = channel,
	file = fileContent['file'],
	title = f'{name} Mistake Report'
	)
	#print(f'Mistake report sent to {name}')

# open workbook and assign sheets
mistakeReportFile = [f for f in os.listdir("Mistake Report XLSX/") if f.endswith('.xlsx')] # lists any xlsx in directory
managementFile = [f for f in os.listdir('./') if f.endswith('.xlsx')] # lists any xlsx in cwd 

try:
	mistakeReportWorkbook = openpyxl.load_workbook('Mistake Report XLSX/' + mistakeReportFile[0])
	managementWorkbook = openpyxl.load_workbook(managementFile[0])
except IndexError:
	print("ERROR: Place mistake report in the folder where this program is located and restart program.")
	sys.exit()

mistakeReportWorkbookNames = mistakeReportWorkbook.sheetnames
managementWorkbookNames = managementWorkbook.sheetnames
sheet1, sheet2 = mistakeReportWorkbook[mistakeReportWorkbookNames[0]], mistakeReportWorkbook[mistakeReportWorkbookNames[1]]
masterList, specific = managementWorkbook[managementWorkbookNames[0]], managementWorkbook[managementWorkbookNames[1]]
firstShift, secondShift = managementWorkbook[managementWorkbookNames[2]], managementWorkbook[managementWorkbookNames[3]]

slack = SlackClient(masterList['E2'].value)
m = Fore.GREEN + 'M' + Style.RESET_ALL
s = Fore.GREEN + 'S' + Style.RESET_ALL
one = Fore.GREEN + '1' + Style.RESET_ALL
two = Fore.GREEN + '2' + Style.RESET_ALL
print('\nSelect from Menu:')
print(f'\t<{m}>-Send All Mistakes')
print(f'\t<{s}>-Send to Specific Loggers')
print(f'\t<{one}>-Send Mistakes-1st Shift')
print(f'\t<{two}>-Send Mistakes-2nd Shift\n')
ans = input('ENTER >>>')
if ans.lower() == 'm':
	managementList = masterList
elif ans.lower() == 's':
	managementList = specific
elif ans == '1':
	managementList = firstShift
elif ans == '2':
	managementList = secondShift
print('Connecting to Slack...')

# finds the amount of rows in management mistakeReportWorkbook (test sheet is masterList)
amountOfRows = 0
for i in range(0, 1048576): # calculates how many rows of data in column 1
	logger = managementList.cell(row=i + 2, column=1).value # goes through each cell in column 1
	amountOfRows += 1
	if logger == None:
		amountOfRows += 1
		break

if amountOfRows == 0:
	print(Fore.RED + 'ERROR: List is empty. Please restart program with Name|MemberID in management.xlsx')

# finds the users channel by using their memberId in management mistakeReportWorkbook (masterList)
channelList = []
for i in range(0, amountOfRows):
	channel = None
	loggerName = managementList.cell(row=i + 2, column=1).value
	loggerMemId = managementList.cell(row=i + 2, column=2).value
	try:
		channel = slack.api_call('im.open', user = loggerMemId)['channel']['id']
		channelList.append({'Name': loggerName, 'Channel': channel})
	except KeyError: # skips channel
		continue

if channelList == []:
	print(Fore.RED + 'ERROR: Please restart program with Active Token in management.xlsx')
	sys.exit()
else:
	print('Slack Token: ' + Fore.GREEN + 'Granted')

# collects a list of the name on the mistake report, will find out how to perform code block via .max_row method
numRowsMistakeNames = 0
for i in range(0, 1048576): # calculates how many rows of data in column 1
	logger = sheet2.cell(row=i + 3, column=1).value # goes through each cell in column 1
	if (logger != None):
		numRowsMistakeNames += 1
		continue
	break

# creates a sorted(set()) list of logger names that have mistakes, located in {loggerName}sheet (A**x T***a)
mistakeNames = sorted((list(set([sheet2.cell(row=i + 3, column=1).value for i in range(numRowsMistakeNames)]))))

# checks to see if any loggers with a channel have a mistake
# paper counts how many pieces of paper were saved
paper = 0
timerStart = time.time()
print('Sending Mistake Reports...')
for user in channelList: 
	for name in mistakeNames:
		if name in user['Name']:
			# please reorder this to reflect the excel doc, lazy
			mistakes = [sheet2.cell(row=i + 3, column=4).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']]

			incidentDate = [sheet2.cell(row=i + 3, column=3).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']]

			enteredDate = [sheet2.cell(row=i + 3, column=2).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']]   

			suite = [sheet2.cell(row=i + 3, column=5).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']] 

			pkgId = [sheet2.cell(row=i + 3, column=6).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']]  

			incidentNotes = [sheet2.cell(row=i + 3, column=8).value 
						for i in range(numRowsMistakeNames) 
						if sheet2.cell(row=i + 3, column=1).value == user['Name']]    

			#creates new mistake workbook
			# this work book is going to be sent to the logger
			mistakeReportWorkbook = openpyxl.Workbook()
			ws = mistakeReportWorkbook.active

			# sets up header to reflect master mistake report
			ws['A1'].value = 'Employee'
			ws['B1'].value = 'Entered Date'
			ws['C1'].value = 'Incident Date'
			ws['D1'].value = 'Mistake Type'
			ws['E1'].value = 'Suite'
			ws['F1'].value = 'Pkg Id'
			ws['G1'].value = 'Incident Notes'

			# formats header to reflect master mistake report
			boldFont = openpyxl.styles.Font(bold = True)
			center = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
			for cell in ws["1:1"]:
				cell.font = boldFont
				cell.alignment = center
			ws['A2'].value = name

			# write information for mistake into spreadsheet
			for i in range(len(mistakes)):
				ws.cell(row=i + 2, column=2).value = enteredDate[i]
				ws.cell(row=i + 2, column=3).value = incidentDate[i]
				ws.cell(row=i + 2, column=4).value = mistakes[i]
				ws.cell(row=i + 2, column=5).value = suite[i]
				ws.cell(row=i + 2, column=6).value = pkgId[i]
				ws.cell(row=i + 2, column=7).value = incidentNotes[i]

			# format date
			for i in range(len(mistakes)):
				cell = ws.cell(row=i + 2, column=2)
				cell.value = cell.value.strftime('%m/%d/%y')
				cell = ws.cell(row=i + 2, column=3)
				cell.value = cell.value.strftime('%m/%d/%y')

			# wrap text = True
			for row in ws.iter_rows():
				for cell in row:
					cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

			# column width
			columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
			for col in columns:
				if col == 'A':
					ws.column_dimensions[col].width = 19.29
					continue
				elif col == 'B':
					ws.column_dimensions[col].width = 13.57
					continue
				elif col == 'C':
					ws.column_dimensions[col].width = 13
					continue
				elif col == 'D':
					ws.column_dimensions[col].width = 48
					continue
				elif col == 'E':
					ws.column_dimensions[col].width = 9.86
					continue
				elif col == 'F':
					ws.column_dimensions[col].width = 12
					continue 
				elif col == 'G':
					ws.column_dimensions[col].width = 78.43
					continue

			mistakeReportWorkbook.save(f'Sent/{name}MistakeReport.xlsx')

			# excel is saved as png and is then sent to logger
			excel2img.export_img(f'Sent/{name}MistakeReport.xlsx', f"Sent/{name}MistakeReport.png", "Sheet", None)

			sendMistakeReport(name, user['Channel'])

			paper += 1
timerEnd = time.time()
print(timerEnd - timerStart)
print(f'You saved {paper} pieces of paper today')
sys.exit()