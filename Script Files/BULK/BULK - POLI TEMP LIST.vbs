'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - POLI TEMP.vbs" 'BULK script that creates a spreadsheet of the POLI/TEMP topics, sections, and revision dates'
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Required for statistical purposes==========================================================================================
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 23			 'manual run time in seconds
STATS_denomination = "I"		 'I is for item
'END OF stats block==============================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""				'Connects to BlueZone
back_to_self				'navigates back to the self screen since POLI/TEMP is super picky
Call check_for_MAXIS(False)	'stops script if user is passworded out

Call navigate_to_MAXIS_screen("POLI", "____")
EMWriteScreen "TEMP", 5, 40
EMWriteScreen "TABLE", 21, 71
transmit

'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and phone numbers
objExcel.cells(1, 1).Value = "TITLE"
objExcel.Cells(1, 2).Value = "SECTION"
objExcel.Cells(1, 3).Value = "REVISED"

FOR i = 1 to 3		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Excel_row = 2 'Declaring the row to start with

'DO...LOOP adds the POLI/TEMP info to the spreadsheet and checks for the end of the page
Do
	MAXIS_row = 6	'Setting or resetting this to look at the top of the list
	DO	'All of this loops until MAXIS_row = 19
		'Reading POLI TEMP info
		EMReadScreen title_info, 45, MAXIS_row, 8
		EMReadScreen section_info, 11, MAXIS_row, 54
		EMReadScreen revised_info, 7, MAXIS_row, 74
		'Adding the case to Excel
		ObjExcel.Cells(excel_row, 1).Value = trim(title_info)
		ObjExcel.Cells(excel_row, 2).Value = trim(section_info)
		ObjExcel.Cells(excel_row, 3).Value = trim(revised_info)
		STATS_counter = STATS_counter + 1								'adds one instance to the stats counter
		If trim(title_info) = "TESTING UPLOAD PROCES" then exit do		'this is the last entry of POLI/TEMP, no page breaks
		excel_row = excel_row + 1										'shifting to the next excel row
		MAXIS_row = MAXIS_row + 1										'
	Loop until MAXIS_row = 21		'Last row on POLI/TEMP screen
	'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s
	PF8
Loop until trim(title_info) = "TESTING UPLOAD PROCES"

'Deleting the last line of POLI/TEMP to clean up the spreadsheet
SET objRange = objExcel.Cells(excel_row, 1).EntireRow
objRange.Delete

'Formatting the columns to auto-fit after they are all finished being created.
FOR i = 1 to 3									'formatting the cells
 	objExcel.Cells(1, i).Font.Bold = True		'bold font
 	objExcel.Columns(i).AutoFit()				'sizing the columns
 NEXT

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

script_end_procedure("Success! The list of current POLI/TEMP topics is now complete.")
