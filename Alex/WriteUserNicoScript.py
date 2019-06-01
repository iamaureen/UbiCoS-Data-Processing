"""
Written by Alex Blackson
Last Modified: May 22, 2019
Instructions for Use: 
1. Run UserNicoSync.py to merge User and Nico dialogues
2. Use output xlsx file from UserNcioSync script as the mergedFile in main()
3. Set destFile in main to destination on local machine
4. Set minTime and maxTime to time range looking to analyze
5. Run script with the command - python WriteUserNicoScript.py
6. Open Excel spreadsheet in specified location
7. Adjust column widths and set columns to "Wrap Text"
8. If you wish, you can save spreadsheet as PDF for better readability 
"""

import pandas as pd
from datetime import datetime
import xlsxwriter


def genDictionary(dirty, minTime, maxTime):
	dirtyRowCount = dirty.shape[0] - 2
	dirtyIndex = 1
	dialogueDictionary  = {}
	while dirtyIndex < dirtyRowCount:

		if isinstance(dirty.iloc[dirtyIndex]['DateTime'], str):
			currDatetime = datetime.strptime(dirty.iloc[dirtyIndex]['DateTime'],'%m/%d/%Y %I:%M:%S %p')

			if currDatetime >= minTime and currDatetime <= maxTime and isinstance(dirty.iloc[dirtyIndex]['Transcript'], str): 
				currUser = dirty.iloc[dirtyIndex]['UserID']
				currOwner = dirty.iloc[dirtyIndex]['Owner']
				currTranscript = dirty.iloc[dirtyIndex]['Transcript']
				currProblem = dirty.iloc[dirtyIndex]['ProblemID']
				currTime = currDatetime.strftime('%H:%M:%S')

				if currUser not in dialogueDictionary:
					dialogueDictionary[currUser] = [(int(currProblem), currTime, currOwner, currTranscript)]
				else:
					dialogueDictionary[currUser].append((int(currProblem), currTime, currOwner, currTranscript))

		dirtyIndex += 1

	return dialogueDictionary

def writeScript(dialogueDictionary, destFile):
	writer = pd.ExcelWriter(destFile, engine='xlsxwriter')
	workbook = writer.book
	workbook.formats[0].set_border_color('white')
	worksheet = workbook.add_worksheet('Dialogue')

	header_format = workbook.add_format({'bold': True, 'italic': True, 'font_size': 16})
	speaker_format = workbook.add_format({'bold': True, 'valign': 'top'})
	time_format = workbook.add_format({'italic': True, 'valign': 'top'})
	problem_format = workbook.add_format({'underline': True, 'bold': True, 'font_size': 14})
	transcript_format = workbook.add_format({'valign': 'top'})

	rowNum = 0

	for user, dialogue in dialogueDictionary.items():
		currProb = -1
		worksheet.write(rowNum, 0, user.upper(), header_format)
		rowNum += 1
		for pair in dialogue:
			if currProb != pair[0]:
				rowNum += 1
				worksheet.write(rowNum, 0, 'Problem ' + str(pair[0]), problem_format)
				rowNum += 1
			worksheet.write(rowNum, 0, pair[1], time_format)
			worksheet.write(rowNum, 1, pair[2] + ':', speaker_format)
			worksheet.write(rowNum, 2, pair[3], transcript_format)
			rowNum += 1
			currProb = pair[0]
		rowNum += 1

	writer.save()

def cleanLogs(mergedFile, destFile, minTime, maxTime):
	dirty = pd.read_csv(mergedFile, names=["StateKey","UserID","DateTime","SessionID","ProblemID","StepID","Owner","DialogueAct","DialogueActConfidence","Spoke","StepAnswer","ClickStep","NicoMove","Answered","Transcript"], infer_datetime_format=True)
	dialogueDictionary = genDictionary(dirty, minTime, maxTime)

	writeScript(dialogueDictionary, destFile)




def main():
	# ADJUST THESE TO YOUR APPROPRIATE LOCAL SOURCE/DESTINATION
	mergedFile = "C:\\Users\\arb17\\Documents\\Research\\Cobi Session 2 Transcripts\\all_dialogue_april19.csv"
	destFile = "C:\\Users\\arb17\\Documents\\Research\\Cobi Session 2 Transcripts\\all_dialogue_march19_clean.xlsx"

	# ADJUST THESE TO YOUR APPROPRIATE DATE RANGE OF STUDY TO ANALYZE
	minTime = datetime(year=2019, month=3, day=4, hour=12, minute=0, second=0)
	maxTime = datetime(year=2019, month=3, day=4, hour=16, minute=0, second=0)
	cleanLogs(mergedFile, destFile, minTime, maxTime)


if __name__ == '__main__': main()