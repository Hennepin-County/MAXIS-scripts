'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - INTERVIEW TEAM CASES WORKLIST.vbs"
start_time = timer
STATS_counter = 0			     'sets the stats counter at one
STATS_manualtime = 	90			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/22/2025", "Worklist can be created for cases interviewed the current day.##~## ##~##If the same day selection is made the file will be saved with a number at the end and there will be multiple worklists for the interview day(s) selected.##~##", "Casey Love, Hennepin County")
call changelog_update("01/13/2025", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================

set basket_detail = CreateObject("Scripting.Dictionary")

'Team 1 Clifton			'Team 2 Coenen			'Team 3 Garrett			'Team 4 Groves
basket_detail.add "X127EQ9", "Adults"			'OLD BASKET STRUCTURE???
basket_detail.add "X127EK8", "Adults" ' - Pending 1"
basket_detail.add "X127EH1", "Adults" ' - Pending 1"
basket_detail.add "X127EP1", "Adults" ' - Pending 1"
basket_detail.add "X127EP2", "Adults" ' - Pending 2"
basket_detail.add "X127EH8", "Adults" ' - Pending 2"
basket_detail.add "X127EP6", "Adults" ' - Pending 2"
basket_detail.add "X127EP7", "Adults" ' - Pending 3"
basket_detail.add "X127EP8", "Adults" ' - Pending 3"
basket_detail.add "X127EP3", "Adults" ' - Pending 3"
basket_detail.add "X127EH7", "Adults" ' - Pending 4"
basket_detail.add "X127EK3", "Adults" ' - Pending 4"
basket_detail.add "X127EK7", "Adults" ' - Pending 4"
basket_detail.add "X127EQ5", "Adults" ' Active 1"
basket_detail.add "X127EQ6", "Adults" ' Active 1"
basket_detail.add "X127EQ7", "Adults" ' Active 1"
basket_detail.add "X127EQ8", "Adults" ' Active 1"
basket_detail.add "X127EX1", "Adults" ' Active 1"
basket_detail.add "X127EX2", "Adults" ' Active 1"
basket_detail.add "X127EX3", "Adults" ' Active 1"
basket_detail.add "X127EX4", "Adults" ' Active 1"
basket_detail.add "X127EX5", "Adults" ' Active 1"
basket_detail.add "X127EX7", "Adults" ' Active 1"
' basket_detail.add "X127F3H", "Adults" ' Active 1"		'DELETE?
basket_detail.add "X127EL7", "Adults" ' Active 2"
basket_detail.add "X127EL8", "Adults" ' Active 2"
basket_detail.add "X127EL9", "Adults" ' Active 2"
basket_detail.add "X127EN1", "Adults" ' Active 2"
basket_detail.add "X127EN2", "Adults" ' Active 2"
basket_detail.add "X127EN3", "Adults" ' Active 2"
basket_detail.add "X127EN5", "Adults" ' Active 2"
basket_detail.add "X127EN4", "Adults" ' Active 2"
basket_detail.add "X127EN7", "Adults" ' Active 2"
basket_detail.add "X127EN8", "Adults" ' Active 3"
basket_detail.add "X127EN9", "Adults" ' Active 3"
basket_detail.add "X127EQ1", "Adults" ' Active 3"
basket_detail.add "X127EQ2", "Adults" ' Active 3"
basket_detail.add "X127EQ3", "Adults" ' Active 3"
basket_detail.add "X127EQ4", "Adults" ' Active 3"
basket_detail.add "X127EX8", "Adults" ' Active 3"
basket_detail.add "X127EX9", "Adults" ' Active 3"
basket_detail.add "X127EG4", "Adults" ' Active 3"
basket_detail.add "X127ED8", "Adults" ' Active 4"
basket_detail.add "X127EE1", "Adults" ' Active 4"
basket_detail.add "X127EE2", "Adults" ' Active 4"
basket_detail.add "X127EE3", "Adults" ' Active 4"
basket_detail.add "X127EE4", "Adults" ' Active 4"
basket_detail.add "X127EE5", "Adults" ' Active 4"
basket_detail.add "X127EE6", "Adults" ' Active 4"
basket_detail.add "X127EE7", "Adults" ' Active 4"
basket_detail.add "X127EL1", "Adults" ' Active 4"
basket_detail.add "X127EL2", "Adults" ' Active 4"
basket_detail.add "X127EL3", "Adults" ' Active 4"
basket_detail.add "X127EL4", "Adults" ' Active 4"
basket_detail.add "X127EL5", "Adults" ' Active 4"
basket_detail.add "X127EL6", "Adults" ' Active 4"

basket_detail.add "X127ET5", "Families" 		'Active 1"
basket_detail.add "X127ET6", "Families" 		'Active 1"
basket_detail.add "X127ET7", "Families" 		'Active 1"
basket_detail.add "X127ET8", "Families" 		'Active 1"
basket_detail.add "X127ET9", "Families" 		'Active 1"
basket_detail.add "X127EZ1", "Families" 		'Active 1"
basket_detail.add "X127ES1", "Families" 		'Active 2"
basket_detail.add "X127ES2", "Families" 		'Active 2"
basket_detail.add "X127ET1", "Families" 		'Active 2"
' basket_detail.add "X127F4E", "Families" 		'Active 2"		'DELETE?
basket_detail.add "X127EZ7", "Families" 		'Active 2"
basket_detail.add "X127FB7", "Families" 		'Active 2"
basket_detail.add "X127ET2", "Families" 		'Active 3"
basket_detail.add "X127ET3", "Families" 		'Active 3"
basket_detail.add "X127ET4", "Families" 		'Active 3"
basket_detail.add "X127ES3", "Families" 		'Active 4"
basket_detail.add "X127ES4", "Families" 		'Active 4"
basket_detail.add "X127ES5", "Families" 		'Active 4"
basket_detail.add "X127ES6", "Families" 		'Active 4"
basket_detail.add "X127ES7", "Families" 		'Active 4"
basket_detail.add "X127ES8", "Families" 		'Active 4"
basket_detail.add "X127ES9", "Families" 		'Active 4"
basket_detail.add "X127EZ6", "Families" 		'- Pending 1"
basket_detail.add "X127EZ8", "Families" 		'- Pending 1"
basket_detail.add "X127EZ9", "Families" 		'- Pending 2"
basket_detail.add "X127EH4", "Families" 		'- Pending 2"
basket_detail.add "X127EH5", "Families" 		'- Pending 3"
basket_detail.add "X127EH6", "Families" 		'- Pending 3"
basket_detail.add "X127EZ3", "Families" 		'- Pending 4"
basket_detail.add "X127EZ4", "Families" 		'- Pending 4"

basket_detail.add "X127F3P", "Adults"   'MA-EPD Adults Basket
basket_detail.add "X127F3K", "Families"  'MA-EPD FAD Basket
' basket_detail.add "X127F3P", "Families - General"		- this is MAEPD

basket_detail.add "X127FE7", "DWP"
basket_detail.add "X127FE8", "DWP"
basket_detail.add "X127FE9", "DWP"
basket_detail.add "X127EY8", "DWP"
basket_detail.add "X127EY9", "DWP"

basket_detail.add "X127FA5", "YET"
basket_detail.add "X127FA6", "YET"
basket_detail.add "X127FA7", "YET"
basket_detail.add "X127FA8", "YET"
basket_detail.add "X127FB1", "YET"
basket_detail.add "X127FA9", "YET"

basket_detail.add "X127EN6", "TEFRA"
basket_detail.add "X127FG1", "Foster Care / IV-E"
basket_detail.add "X127EW6", "Foster Care / IV-E"
basket_detail.add "X1274EC", "Foster Care / IV-E"
basket_detail.add "X127FG2", "Foster Care / IV-E"
basket_detail.add "X127EW4", "Foster Care / IV-E"

basket_detail.add "X127EM8", "GRH / HS - Adults Pending"
basket_detail.add "X127FE6", "GRH / HS - Adults Pending"
basket_detail.add "X127EZ2", "GRH / HS - Families Pending"
basket_detail.add "X127EM2", "GRH / HS - Maintenance"
basket_detail.add "X127EH9", "GRH / HS - Maintenance"
basket_detail.add "X127EJ4", "GRH / HS - Maintenance"
basket_detail.add "X127EH2", "GRH / HS - Maintenance"
basket_detail.add "X127EP4", "GRH / HS - Maintenance"
basket_detail.add "X127EK5", "GRH / HS - Maintenance"
basket_detail.add "X127EG5", "GRH / HS - Maintenance"

'basket_detail.add "X127EG4", "MIPPA"
basket_detail.add "X127F3D", "MA - BC"

basket_detail.add "X127EF8", "1800"
basket_detail.add "X127EF9", "1800"
basket_detail.add "X127EG9", "1800"
basket_detail.add "X127EG0", "1800"

basket_detail.add "X1275H5", "Privileged Cases"
basket_detail.add "X127FAT", "Privileged Cases"
basket_detail.add "X127F3H", "Privileged Cases"
'Contacted Case Mgt
basket_detail.add "X127FG6", "LTC+"           '"Kristen Kasem"
basket_detail.add "X127FG7", "LTC+"           '"Kristen Kasem"
basket_detail.add "X127EM3", "LTC+"           '"True L. or Gina G."
basket_detail.add "X127EM4", "LTC+"            '"True L. or Gina G."
basket_detail.add "X127EW7", "LTC+"            '"Kimberly Hill"
basket_detail.add "X127EW8", "LTC+"            '"Kimberly Hill"
basket_detail.add "X127FF4", "LTC+"            '"Alyssa Taylor"
basket_detail.add "X127FF5", "LTC+"            '"Alyssa Taylor"
basket_detail.add "X127FF8", "LTC+"				'"Contracted - North Memorial"
basket_detail.add "X127FF6", "LTC+"				'"Contracted - HCMC"
basket_detail.add "X127FF7", "LTC+"				'"Contracted - HCMC"

' basket_detail.add "X127EK4", "LTC+ - General"
' basket_detail.add "X127EK9", "LTC+ - General"
' basket_detail.add "X127EH1", "LTC+"
basket_detail.add "X127EH3", "LTC+"
' basket_detail.add "X127EH4", "LTC+"
' basket_detail.add "X127EH5", "LTC+"
' basket_detail.add "X127EH6", "LTC+"
' basket_detail.add "X127EH7", "LTC+"
basket_detail.add "X127EJ8", "LTC+"
basket_detail.add "X127EK1", "LTC+"
basket_detail.add "X127EK2", "LTC+"
' basket_detail.add "X127EK3", "LTC+"
basket_detail.add "X127EK4", "LTC+"
basket_detail.add "X127EK6", "LTC+"
' basket_detail.add "X127EK7", "LTC+"
' basket_detail.add "X127EK8", "LTC+"
basket_detail.add "X127EK9", "LTC+"
basket_detail.add "X127EM9", "LTC+"
' basket_detail.add "X127EN6", "LTC+"
basket_detail.add "X127EP5", "LTC+"
basket_detail.add "X127EP9", "LTC+"
basket_detail.add "X127EZ5", "LTC+"
basket_detail.add "X127F3F", "LTC+"
basket_detail.add "X127FE5", "LTC+"
basket_detail.add "X127FH4", "LTC+"
basket_detail.add "X127FH5", "LTC+"
basket_detail.add "X127FI2", "LTC+"
basket_detail.add "X127FI7", "LTC+"

basket_detail.add "X127FI1", "METS Retro Request"


curr_month = MonthName(DatePart("M",date))
curr_day = DatePart("d",date)
curr_year = DatePart("yyyy", date)

manager_log_file_path 				= "https://hennepin.sharepoint.com/teams/InterviewPhoneHSRs-Supervisors/Shared%20Documents/Management%20Team-HSR%20Interviews/Manager%20Log%20" & curr_year & "%20" & curr_month & ".xlsx"
staff_assignment_log_file_path 		= "https://hennepin.sharepoint.com/teams/InterviewPhoneHSRs-EligibilityHSRs/Shared%20Documents/Eligibility%20HSRs/Eligibility%20Staff%20" & curr_year & "%20" & curr_month & ".xlsx"

template_manager_log_file_path 			= t_drive & "\Eligibility Support\Assignments\Interview Team Cases Worklists\Support Documents\Template Manager Log.xlsx"
template_staff_assignment_log_file_path = t_drive & "\Eligibility Support\Assignments\Interview Team Cases Worklists\Support Documents\Eligibility Staff Template.xlsx"

end_message = ""
'END DECLARATIONS ==========================================================================================================


'THE SCRIPT ===========================================================================================================

'Now we are ready to start the script

'Dialog is just to alert user that the script starts with information gathering and takes time.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 115, "Interview Team Cases Worklist"
  ButtonGroup ButtonPressed
    OkButton 125, 90, 50, 15
  Text 10, 10, 175, 25, "This script will create a list of cases with interviews completed by the Interview Team. It will also update the files used by the processing team to assign work. "
  Text 10, 45, 165, 20, "First the script needs to gather some information, which may take a minute or two. "
  Text 10, 75, 145, 10, "Please be patient, the script is running!"
  Text 10, 90, 105, 10, "Press OK to start info gather."
EndDialog

'no options on the dialog so no loop is necessary
dialog Dialog1
cancel_confirmation

'FILE MANAGEMENT ===========================================================================================================
'This section ensures the files needed for the run are up available and up to date.

'Sharepoint files do not have the same 'FileExists' possibility so we need to try to open them and see if there is an error to determine if a new file is needed.
If curr_day < 7 Then
	On Error Resume Next			'allow for errors to occur and then use them to identify missing files.

	'not using the function because we need to disable the alerts BEFORE trying to open the workbook
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	Set objWorkbook = objExcel.Workbooks.Open(staff_assignment_log_file_path) 'Opens an excel file from a specific URL
	' MsgBox "Error Number: " & Err.Number
	If Err.Number <> 0 Then
		Set ObjExcel = Nothing
		Set objWorkbook = Nothing

		Call excel_open(template_staff_assignment_log_file_path, False, False, ObjExcel, objWorkbook)
		ObjExcel.ActiveWorkbook.SaveAs staff_assignment_log_file_path
		ObjExcel.ActiveWorkbook.Close
		ObjExcel.Application.Quit
		ObjExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
		end_message = end_message & "New Eligibility Staff Excel created for " & curr_month & ". The file has been saved to the TEAMS location." & vbCr
	End If
	If Err.Number = 0 Then
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
	End If

	Err.Clear

	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = False
	objExcel.DisplayAlerts = False
	Set objWorkbook = objExcel.Workbooks.Open(manager_log_file_path) 'Opens an excel file from a specific URL
	' MsgBox "Error Number: " & Err.Number
	If Err.Number <> 0 Then
		Set ObjExcel = Nothing
		Set objWorkbook = Nothing

		first_of_the_month = DatePart("M",date) & "/1/" & curr_year
		first_of_the_month = DateAdd("d", 0, first_of_the_month)

		Call excel_open(template_manager_log_file_path, False, False, ObjExcel, objWorkbook)
		ObjExcel.worksheets("Counts").Activate
		ObjExcel.Cells(2,3).Value = first_of_the_month	'set up the counts sheet for the current month
		ObjExcel.worksheets("Manager log").Activate

		ObjExcel.ActiveWorkbook.SaveAs manager_log_file_path
		ObjExcel.ActiveWorkbook.Close
		ObjExcel.Application.Quit
		ObjExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
		end_message = end_message & "New Manager Log Excel created for " & curr_month & ". The file has been saved to the TEAMS location." & vbCr
	End If
	'Call script_end_procedure("The Staff assignment log for " & curr_month & " " & curr_year & " is not available. Create the Excel first and Rerun the script")

	If Err.Number = 0 Then
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit

		Set ObjExcel = Nothing
		Set objWorkbook = Nothing
	End If

	Err.Clear

	On Error Goto 0
End If


'Setting the folder paths and objects to handle folder and file manipulation
interview_team_cases_folder = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage"
Set objFolder = objFSO.GetFolder(interview_team_cases_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

interview_team_cases_already_on_worklist = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List"
worklist_folder = t_drive & "\Eligibility Support\Assignments\Interview Team Cases Worklists"

'constants for the array of all dates - this allows selection of cases by date
Const date_const 	= 01
Const month_const 	= 02
Const day_const		= 03
Const count_const	= 04
Const checkbox_const = 05
const added_count	= 06
const fam_count 	= 07
const adul_count	= 08
Const dates_last_const	= 20

'This array gathers all the dates but is not sorted
Dim TEMP_ARRAY()
ReDim TEMP_ARRAY(dates_last_const, 0)

'We will use this array to sort the original array in date order
Dim DATES_WITH_INTERVIEWS_ARRAY()
ReDim DATES_WITH_INTERVIEWS_ARRAY(dates_last_const, 0)

'setting some initial variables for array use
each_month = 0
Total_count = 0
earliest_date = ""

'Looking at each txt file in the assignments folder to capture Expedited Determination information
For Each objFile in colFiles							'looping through each file
	file_created_date = objFile.DateCreated				'Reading the date created
	Total_count = Total_count + 1						'count all cases

	'creating some better formatting of dates (date created parameter has time associated and causes issues)
	intvw_month = DatePart("m", file_created_date)
	intvw_day = DatePart("d", file_created_date)
	intvw_year = DatePart("yyyy", file_created_date)
	file_date = intvw_month & "/" & intvw_day & "/" & intvw_year
	file_date = DateAdd("d", 0, file_date)

	'seeing if we already have this date in our list of dates
	date_found = False
	For chkn_wg = 0 to UBound(TEMP_ARRAY, 2)
		If file_date = TEMP_ARRAY(date_const, chkn_wg) Then
			TEMP_ARRAY(count_const, chkn_wg) = TEMP_ARRAY(count_const, chkn_wg) + 1
			date_found = True
			Exit For
		End If
	Next

	'If the date was not found, we will add it to the list of dates
	If date_found = False Then
		ReDim Preserve TEMP_ARRAY(dates_last_const, each_month)
		TEMP_ARRAY(date_const, 	each_month) = file_date
		TEMP_ARRAY(count_const, each_month) = 1
		TEMP_ARRAY(month_const, each_month) = intvw_month
		TEMP_ARRAY(day_const, 	each_month) = intvw_day

		each_month = each_month + 1

		'Need to define the first date in the list for easy sorting later
		If earliest_date = "" Then
			earliest_date = file_date
		Else
			If DateDiff("d", earliest_date, file_date) < 0 Then earliest_date = file_date
		End If
	End If
Next

'checking to ensure there are some cases to create a worklist with from a day prior to today
If earliest_date = "" Then Call script_end_procedure("There does not appear to be any outstanding cases from interviews on previous days. To check for previous worklists reference the script instructions for the worklist folder.")

'Loop through each day from the earliest day forward to put the dates in order from oldest to newest
date_to_assess = earliest_date
tomorrow = DateAdd("d", 1, date)
cow = 0
Do While DateDiff("d", tomorrow, date_to_assess) <> 0			'we stop when we get to today and DO NOT include today
	For pig = 0 to UBound(TEMP_ARRAY, 2)					'go through the temp array to find a match
		If DateDiff("d", TEMP_ARRAY(date_const, pig), date_to_assess) = 0 Then
			ReDim Preserve DATES_WITH_INTERVIEWS_ARRAY(dates_last_const, cow)					'Add the date with all parameters to the new array
			DATES_WITH_INTERVIEWS_ARRAY(date_const, 	cow) = TEMP_ARRAY(date_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(count_const, 	cow) = TEMP_ARRAY(count_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(month_const, 	cow) = TEMP_ARRAY(month_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(day_const, 		cow) = TEMP_ARRAY(day_const, pig)
			DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, cow) = unchecked
			cow = cow + 1
		End If
	Next
	date_to_assess = DateAdd("d", 1, date_to_assess)		'incrementing the date we are assessing for the loop
Loop

'Now we can display all of the dates with case count in a dialog for selection.
dlg_len = 80
For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
	dlg_len = dlg_len + 15
Next

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 316, dlg_len, "Select Dates for Worklist"
	Text 10, 35, 195, 10, "To create a worklist, select the date(s) to include:"
	Text 10, 10, 250, 20, "All interview team recorded cases have been checked and there appear to be " & Total_count & " cases that have not yet been added to a worklist."
	Text 20, 50, 40, 15, "Check Here to Select"
	Text 80, 60, 25, 10, "Date"
	Text 140, 60, 35, 10, "Count"
	y_pos = 70
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		CheckBox 30, y_pos, 30, 10, DATES_WITH_INTERVIEWS_ARRAY(month_const, chkn_wg) & "/" & DATES_WITH_INTERVIEWS_ARRAY(day_const, chkn_wg), DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg)
		Text 80, y_pos, 50, 10, DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg)
		Text 145, y_pos, 50, 10, DATES_WITH_INTERVIEWS_ARRAY(count_const, chkn_wg)
		y_pos = y_pos + 15
	Next
	ButtonGroup ButtonPressed
		OkButton 200, dlg_len-25, 50, 15
		CancelButton 255, dlg_len-25, 50, 15
EndDialog

Do
	err_msg = ""
	dialog Dialog1
	cancel_confirmation

	'Only requirement is to select at least one date.
	date_selected = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked then date_selected = True
	Next

	If date_selected = False Then err_msg = err_msg & vbCr & "* Select at least 1 date."
	If err_msg <> "" Then MsgBox "* * * NOTICE * * * " & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
Loop until err_msg = ""

developer_mode = False
If user_ID_for_validation = "CALO001" Then
	run_in_dev = MsgBox("Do you want to run in developer mode?", vbQuestion + vbYesNo, "Developer Mode")
	If run_in_dev = vbYes Then developer_mode = True
End If

const x_numb_const		= 00
const case_numb_const	= 01
const population_const	= 02
const intvw_date_const	= 03
const appears_xfs_const	= 04
const cash_select_const	= 05
const cash_prog_const	= 06
const grh_select_const	= 07
const snap_select_const	= 08
const emer_select_const	= 09
const assign_pop_const	= 10
const assigned_to_const	= 11
const mngr_excel_row	= 12
const file_path_const 	= 13
const file_name_const 	= 14
const end_const			= 20

DIM CASES_TO_ASSIGN_ARRAY()
ReDIM CASES_TO_ASSIGN_ARRAY(end_const, 0)
intvws_count = 0
case_count = 0

'creating some objects needed for XML handling
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set xml = CreateObject("Msxml2.DOMDocument")

xml_case_numbs = "*"
'Creating an object for the stream of text which we'll use frequently
Dim objTextStream

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_name = objFile.Name
	file_created_date = objFile.DateCreated					'Reading the date created

	'determining if the XML is for a file in the dates selected by the user
	save_file = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DateDiff("d", DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg), file_created_date) = 0 Then
			If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then save_file = True
		End If
	Next

	'if this is from a date selected, we need to read details of the file for the worklist
	If save_file = True Then
		STATS_counter = STATS_counter + 1
		xmlPath = objFile.Path												'identifying the current file
		With (CreateObject("Scripting.FileSystemObject"))
			Set file_object = CreateObject("Scripting.FileSystemObject")										'Create another FSO
			Set xml_sig_command = file_object.OpenTextFile(xmlPath)			'Open the text file

			name_of_xml = file_object.GetFileName(xmlPath)
			If InStr(name_of_xml, "details") Then
				If xml_sig_command.AtEndOfStream Then
					attachment_here = ""
					Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", "XML in Interview Team Tracking is BLANK", 1, False, "", "", False, "", "XML File Path: " & xmlPath, True, attachment_here, True)
					xml_sig_command.Close
				ElseIf .FileExists(xmlPath) = True then
					xml_sig_command.Close
					xmlDoc.Async = False

					' Load the XML file
					xmlDoc.load(xmlPath)

					'reads data about the case from the XML
					set node = xmlDoc.SelectSingleNode("//CaseBasket")
					case_basket_numb = node.text

					set node = xmlDoc.SelectSingleNode("//CaseNumber")
					MAXIS_case_number = node.text
					xml_case_numbs = xml_case_numbs & trim(MAXIS_case_number) & "*"

					set node = xmlDoc.SelectSingleNode("//InterviewDate")
					interview_date = node.text
					interview_date = DateAdd("d", 0, interview_date)

					set node = xmlDoc.SelectSingleNode("//ScriptRunDate")
					script_date = node.text
					script_date = DateAdd("d", 0, script_date)

					set node = xmlDoc.SelectSingleNode("//CASHRequest")
					req_cash = node.text
					req_cash = req_cash * 1

					cash_type = ""
					If req_cash = True Then
						set node = xmlDoc.SelectSingleNode("//TypeOfCASH")
						cash_type = node.text
					End If

					If DateDiff("d", #1/13/2025#, script_date) > 0 Then
						set node = xmlDoc.SelectSingleNode("//GRHRequest")
						req_grh = node.text
						req_grh = req_grh * 1
					End If

					set node = xmlDoc.SelectSingleNode("//SNAPRequest")
					req_snap = node.text
					req_snap = req_snap * 1

					set node = xmlDoc.SelectSingleNode("//EMERRequest")
					req_emer = node.text
					req_emer = req_emer * 1

					set node = xmlDoc.SelectSingleNode("//ExpeditedDetermination")
					exp_det = node.text
					If exp_det <> "" Then exp_det = exp_det * 1

					case_basket_numb = UCase(case_basket_numb)
					population = ""
					If basket_detail.Exists(case_basket_numb) Then
						population = basket_detail.Item(case_basket_numb)
					Else
						population = "UNKNOWN"
					End If

					population_section = ""
					If population = "Adults" 					Then population_section = "Adults"
					If population = "GRH / HS - Adults Pending" Then population_section = "Adults"
					If population = "GRH / HS - Maintenance" 	Then population_section = "Adults"
					If population = "MA - BC" 					Then population_section = "Adults"
					If population = "1800" 						Then population_section = "Adults"
					If population = "LTC+" 						Then population_section = "Adults"

					If population = "Families" 					Then population_section = "Families"
					If population = "DWP" 						Then population_section = "Families"
					If population = "YET" 						Then population_section = "Families"
					If population = "TEFRA" 					Then population_section = "Families"
					If population = "Foster Care / IV-E" 		Then population_section = "Families"
					If population = "GRH / HS - Families Pending" Then population_section = "Families"

					If population = "METS Retro Request" 		Then population_section = "Families"
					If population = "Privileged Cases" 			Then population_section = "Families"
					If population = "UNKNOWN" 					Then population_section = "Adults"

					ReDim preserve CASES_TO_ASSIGN_ARRAY(end_const, intvws_count)
					CASES_TO_ASSIGN_ARRAY(x_numb_const, intvws_count) 		= case_basket_numb
					CASES_TO_ASSIGN_ARRAY(case_numb_const, intvws_count) 	= MAXIS_case_number
					CASES_TO_ASSIGN_ARRAY(population_const, intvws_count) 	= population
					CASES_TO_ASSIGN_ARRAY(intvw_date_const, intvws_count) 	= interview_date
					If exp_det = True Then CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count) 	= True
					If exp_det <> True Then CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count) 	= False
					CASES_TO_ASSIGN_ARRAY(cash_select_const, intvws_count) 	= req_cash
					If req_cash = True Then CASES_TO_ASSIGN_ARRAY(cash_prog_const, intvws_count) 	= cash_type
					CASES_TO_ASSIGN_ARRAY(grh_select_const, intvws_count) 	= req_grh
					CASES_TO_ASSIGN_ARRAY(snap_select_const, intvws_count) 	= req_snap
					CASES_TO_ASSIGN_ARRAY(emer_select_const, intvws_count) 	= req_emer
					CASES_TO_ASSIGN_ARRAY(assign_pop_const, intvws_count) 	= population_section
					CASES_TO_ASSIGN_ARRAY(assigned_to_const, intvws_count) 	= ""
					CASES_TO_ASSIGN_ARRAY(file_path_const, intvws_count) 	= xmlPath
					CASES_TO_ASSIGN_ARRAY(file_name_const, intvws_count) 	= file_name

					'TODO - If we want to add information to the XML, we need to read it all and then rewrite it all to the xml doc - when we save it, the file will be overwritten

					case_count = case_count + 1
					intvws_count = intvws_count + 1

				End If
			End If
		End With
	End If
Next

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_name = objFile.Name
	file_created_date = objFile.DateCreated					'Reading the date created

	'determining if the XML is for a file in the dates selected by the user
	save_file = False
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DateDiff("d", DATES_WITH_INTERVIEWS_ARRAY(date_const, chkn_wg), file_created_date) = 0 Then
			If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then save_file = True
		End If
	Next

	'if this is from a date selected, we need to read details of the file for the worklist
	If save_file = True Then
		STATS_counter = STATS_counter + 1
		xmlPath = objFile.Path												'identifying the current file
		With (CreateObject("Scripting.FileSystemObject"))
			' 'Creating an object for the stream of text which we'll use frequently
			' Dim objTextStream

			Set file_object = CreateObject("Scripting.FileSystemObject")										'Create another FSO
			Set xml_sig_command = file_object.OpenTextFile(xmlPath)			'Open the text file

			name_of_xml = file_object.GetFileName(xmlPath)
			If InStr(name_of_xml, "started") Then
				If xml_sig_command.AtEndOfStream Then
					attachment_here = ""
					Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", "XML in Interview Team Tracking is BLANK", 1, False, "", "", False, "", "XML File Path: " & xmlPath, True, attachment_here, True)
					xml_sig_command.Close
				ElseIf .FileExists(xmlPath) = True then
					xml_sig_command.Close
					xmlDoc.Async = False

					' Load the XML file
					xmlDoc.load(xmlPath)

					'reads data about the case from the XML
					set node = xmlDoc.SelectSingleNode("//CaseBasket")
					case_basket_numb = node.text

					set node = xmlDoc.SelectSingleNode("//CaseNumber")
					MAXIS_case_number = node.text
					search_numb = "*" & trim(MAXIS_case_number) & "*"

					set node = xmlDoc.SelectSingleNode("//ScriptRunDate")
					script_date = node.text
					script_date = DateAdd("d", 0, script_date)


					case_basket_numb = UCase(case_basket_numb)
					population = ""
					If basket_detail.Exists(case_basket_numb) Then
						population = basket_detail.Item(case_basket_numb)
					Else
						population = "UNKNOWN"
					End If

					population_section = ""
					If population = "Adults" 					Then population_section = "Adults"
					If population = "GRH / HS - Adults Pending" Then population_section = "Adults"
					If population = "GRH / HS - Maintenance" 	Then population_section = "Adults"
					If population = "MA - BC" 					Then population_section = "Adults"
					If population = "1800" 						Then population_section = "Adults"
					If population = "LTC+" 						Then population_section = "Adults"

					If population = "Families" 					Then population_section = "Families"
					If population = "DWP" 						Then population_section = "Families"
					If population = "YET" 						Then population_section = "Families"
					If population = "TEFRA" 					Then population_section = "Families"
					If population = "Foster Care / IV-E" 		Then population_section = "Families"
					If population = "GRH / HS - Families Pending" Then population_section = "Families"

					If population = "METS Retro Request" 		Then population_section = "Families"
					If population = "Privileged Cases" 			Then population_section = "Families"
					If population = "UNKNOWN" 					Then population_section = "Adults"

					If InStr(xml_case_numbs, search_numb) = 0 Then
						ReDim preserve CASES_TO_ASSIGN_ARRAY(end_const, intvws_count)
						CASES_TO_ASSIGN_ARRAY(x_numb_const, intvws_count) 		= case_basket_numb
						CASES_TO_ASSIGN_ARRAY(case_numb_const, intvws_count) 	= MAXIS_case_number
						CASES_TO_ASSIGN_ARRAY(population_const, intvws_count) 	= population
						CASES_TO_ASSIGN_ARRAY(intvw_date_const, intvws_count) 	= script_date
						CASES_TO_ASSIGN_ARRAY(appears_xfs_const, intvws_count)	= "?"
						CASES_TO_ASSIGN_ARRAY(assign_pop_const, intvws_count) 	= population_section
						CASES_TO_ASSIGN_ARRAY(assigned_to_const, intvws_count) 	= ""
						CASES_TO_ASSIGN_ARRAY(file_path_const, intvws_count) 	= xmlPath
						CASES_TO_ASSIGN_ARRAY(file_name_const, intvws_count) 	= file_name

						'TODO - If we want to add information to the XML, we need to read it all and then rewrite it all to the xml doc - when we save it, the file will be overwritten

						case_count = case_count + 1
						intvws_count = intvws_count + 1
					ElseIf developer_mode = False Then
						'moving each file to the folder for cases already in a worklist
						.MoveFile xmlPath , interview_team_cases_already_on_worklist & "\" & file_name
						' .DeleteFile(xmlPath)
					End If
				End If
			End If
		End With
	End If
Next
set xmlDoc = nothing

end_message = end_message & vbCr & "Interviews from " & list_of_dates & " have been added to a worklist."
end_message = end_message & vbCr & "Cases found with an interview completed that needs processing: " & case_count
If developer_mode = False Then
	end_message = end_message & vbCr & vbCr & "Worklist has been left open and can be found here:"
	end_message = end_message & vbCr & full_worklist_file_name
Else
	end_message = end_message & vbCr & "DEVELOPER MODE - FILE NOT SAVED"
End If

'TODO - expand these columns to copy completion data BACK to the manager log - need to identify more columns that have the worker enterd processing notes and dates
mnger_logs_case_numb_col 	= 1
mnger_logs_date_assign_col 	= 2
mnger_logs_appears_xfs_col 	= 3
mnger_log_area_col 			= 4
mnger_log_processor_col		= 5

assign_logs_case_numb_col 		= 1
assign_logs_date_assign_col 	= 2
assign_logs_appears_xfs_col 	= 3
assign_log_area_col 			= 4
assign_log_processor_col		= 5

If developer_mode = False Then
	Call excel_open(manager_log_file_path, True, False, ObjMngrExcel, objMngrWorkbook)
	ObjMngrExcel.worksheets("Manager log").Activate
	excel_row = 1
	Do
		excel_row = excel_row + 1
		list_case_numb = trim(ObjMngrExcel.Cells(excel_row, mnger_logs_case_numb_col).Value)
	Loop until list_case_numb = ""


	total_cases_assigned_count = 0
	fam_cases_assigned_count = 0
	adul_cases_assigned_count = 0

	start_excel = excel_row
	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		ObjMngrExcel.Cells(excel_row, mnger_log_area_col).Value			= ""
		ObjMngrExcel.Cells(excel_row, mnger_logs_case_numb_col).Value  	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
		' ObjMngrExcel.Cells(excel_row, mnger_logs_date_assign_col).Value 	= date
		ObjMngrExcel.Cells(excel_row, mnger_logs_date_assign_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)

		If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then ObjMngrExcel.Cells(excel_row, mnger_logs_appears_xfs_col).Value = "Yes"
		If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = "?" Then ObjMngrExcel.Cells(excel_row, mnger_logs_appears_xfs_col).Value = "?"
		ObjMngrExcel.Cells(excel_row, mnger_log_area_col).Value  		= CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow)
		CASES_TO_ASSIGN_ARRAY(mngr_excel_row, cow) = excel_row

		total_cases_assigned_count = total_cases_assigned_count + 1
		If CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow) = "Families" Then fam_cases_assigned_count = fam_cases_assigned_count + 1
		If CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow) = "Adults" Then adul_cases_assigned_count = adul_cases_assigned_count + 1

		excel_row = excel_row + 1
	Next
	end_excel = excel_row

	' For xl_row = start_excel to end_excel
	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		If trim(ObjMngrExcel.Cells(CASES_TO_ASSIGN_ARRAY(mngr_excel_row, cow), mnger_log_processor_col).Value) = "" Then
			ObjMngrExcel.Cells(CASES_TO_ASSIGN_ARRAY(mngr_excel_row, cow), mnger_log_processor_col).Value = "=IFERROR(XLOOKUP(1, (T_Processors[Area  (SORT-4)]=[@Area])*(T_Processors[Match HSR '#]=[@[Match HSR]]), T_Processors[Processors  (SORT-1)]),  " & Chr(34) & Chr(34) & ")"
		End If
	Next

	'We loop back through the Excel so that it has time to generate the processor name.
	'When done in the initial loop, some of the names were missed.
	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		CASES_TO_ASSIGN_ARRAY(assigned_to_const, cow) = trim(ObjMngrExcel.Cells(CASES_TO_ASSIGN_ARRAY(mngr_excel_row, cow), mnger_log_processor_col).Value)
	Next

	Call excel_open(staff_assignment_log_file_path, True, False, ObjStaffExcel, objStaffWorkbook)

	ObjStaffExcel.worksheets("Families").Activate

	excel_row = 1
	Do
		excel_row = excel_row + 1
		list_case_numb = trim(ObjStaffExcel.Cells(excel_row, assign_logs_case_numb_col).Value)
	Loop until list_case_numb = ""

	unlock_end = excel_row + fam_cases_assigned_count
	ObjStaffExcel.ActiveWorkbook.ActiveSheet.Unprotect
	' ObjStaffExcel.ActiveWorkbook.ActiveSheet.Range("A3:E" & unlock_end).Locked = False

	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		If CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow) = "Families" Then
			ObjStaffExcel.Cells(excel_row, assign_logs_case_numb_col).Value 	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
			ObjStaffExcel.Cells(excel_row, assign_logs_date_assign_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
			If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then ObjStaffExcel.Cells(excel_row, assign_logs_appears_xfs_col).Value = "Yes"
			If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = "?" Then ObjMngrExcel.Cells(excel_row, assign_logs_appears_xfs_col).Value = "?"
			ObjStaffExcel.Cells(excel_row, assign_log_area_col).Value 		= CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow)
			ObjStaffExcel.Cells(excel_row, assign_log_processor_col).Value	= CASES_TO_ASSIGN_ARRAY(assigned_to_const, cow)

			excel_row = excel_row + 1
		End If
	Next

	' ObjStaffExcel.ActiveWorkbook.ActiveSheet.Range("A3:E" & unlock_end).Locked = True
	ObjStaffExcel.ActiveWorkbook.ActiveSheet.Protect

	ObjStaffExcel.worksheets("Adults").Activate

	excel_row = 1
	Do
		excel_row = excel_row + 1
		list_case_numb = trim(ObjStaffExcel.Cells(excel_row, assign_logs_case_numb_col).Value)
	Loop until list_case_numb = ""

	unlock_end = excel_row + adul_cases_assigned_count
	ObjStaffExcel.ActiveWorkbook.ActiveSheet.Unprotect
	' ObjStaffExcel.ActiveWorkbook.ActiveSheet.Range("A3:E" & unlock_end).Locked = False

	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		If CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow) = "Adults" Then
			ObjStaffExcel.Cells(excel_row, assign_logs_case_numb_col).Value 	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
			ObjStaffExcel.Cells(excel_row, assign_logs_date_assign_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
			If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then ObjStaffExcel.Cells(excel_row, assign_logs_appears_xfs_col).Value = "Yes"
			If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = "?" Then ObjMngrExcel.Cells(excel_row, assign_logs_appears_xfs_col).Value = "?"
			ObjStaffExcel.Cells(excel_row, assign_log_area_col).Value 		= CASES_TO_ASSIGN_ARRAY(assign_pop_const, cow)
			ObjStaffExcel.Cells(excel_row, assign_log_processor_col).Value	= CASES_TO_ASSIGN_ARRAY(assigned_to_const, cow)

			excel_row = excel_row + 1
		End If
	Next
	'ADD an LOCK to the spreadsheet columns with information
	' ObjStaffExcel.ActiveWorkbook.ActiveSheet.Range("A3:E" & unlock_end).Locked = True
	ObjStaffExcel.ActiveWorkbook.ActiveSheet.Protect

	'save the manager log file and close
	' objStaffWorkbook.Save()
	' ObjStaffExcel.ActiveWorkbook.Close
	' ObjStaffExcel.Application.Quit
	' ObjStaffExcel.Quit
End If


'Here is the worklist creation section
'these constants are to document the columns  - using a constant supports future changes
Const basket_col		= 1
Const case_numb_col 	= 2
Const population_col 	= 3
Const intvw_date_col 	= 4
Const exp_det_col 		= 5
Const assignment_col	= 6
Const cash_col 			= 7
Const cash_type_col		= 8
Const grh_col 			= 9
Const snap_col 			= 10
Const emer_col 			= 11

'Creating the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first row with header information
ObjExcel.Cells(1, basket_col).Value = "CASELOAD"
ObjExcel.Cells(1, case_numb_col).Value = "CASE NUMBER"
ObjExcel.Cells(1, population_col).Value = "POPULATION"
ObjExcel.Cells(1, intvw_date_col).Value = "INTERVIEW COMPLETED"
ObjExcel.Cells(1, exp_det_col).Value = "APPEARS EXPEDITED"
ObjExcel.Cells(1, assignment_col).Value = "ASSIGNED TO"

ObjExcel.Cells(1, cash_col).Value = "CASH"
ObjExcel.Cells(1, cash_type_col).Value = "CASH TYPE"
ObjExcel.Cells(1, grh_col).Value = "GRH"
ObjExcel.Cells(1, snap_col).Value = "SNAP"
ObjExcel.Cells(1, emer_col).Value = "EMER"
FOR cow = 1 to emer_col							'formatting the cells'
	objExcel.Cells(1, cow).Font.Bold = True		'bold font'
NEXT

Set mvFSO = CreateObject("Scripting.FileSystemObject")

excel_row = 2
For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
	If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "details") Then
		'Add the file information to the Excel document for the worklist
		ObjExcel.Cells(excel_row, basket_col).Value 		= CASES_TO_ASSIGN_ARRAY(x_numb_const, cow)
		ObjExcel.Cells(excel_row, case_numb_col).Value 		= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
		ObjExcel.Cells(excel_row, population_col).Value 	= CASES_TO_ASSIGN_ARRAY(population_const, cow)
		ObjExcel.Cells(excel_row, intvw_date_col).Value 	= CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
		If CASES_TO_ASSIGN_ARRAY(appears_xfs_const, cow) = True Then ObjExcel.Cells(excel_row, exp_det_col).Value = "Yes"
		If CASES_TO_ASSIGN_ARRAY(cash_select_const, cow) = True Then
			ObjExcel.Cells(excel_row, cash_col).Value = "True"
			ObjExcel.Cells(excel_row, cash_type_col).Value = CASES_TO_ASSIGN_ARRAY(cash_prog_const, cow)
		End If
		If CASES_TO_ASSIGN_ARRAY(grh_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, grh_col).Value = "True"
		If CASES_TO_ASSIGN_ARRAY(snap_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, snap_col).Value = "True"
		If CASES_TO_ASSIGN_ARRAY(emer_select_const, cow) 	= True Then ObjExcel.Cells(excel_row, emer_col).Value = "True"
	End If

	If InStr(CASES_TO_ASSIGN_ARRAY(file_path_const, cow), "started") Then
		'Add the file information to the Excel document for the worklist
		ObjExcel.Cells(excel_row, basket_col).Value 	= CASES_TO_ASSIGN_ARRAY(x_numb_const, cow)
		ObjExcel.Cells(excel_row, case_numb_col).Value 	= CASES_TO_ASSIGN_ARRAY(case_numb_const, cow)
		ObjExcel.Cells(excel_row, population_col).Value = CASES_TO_ASSIGN_ARRAY(population_const, cow)
		ObjExcel.Cells(excel_row, intvw_date_col).Value = CASES_TO_ASSIGN_ARRAY(intvw_date_const, cow)
		ObjExcel.Cells(excel_row, exp_det_col).Value = "?"
	End If
	excel_row = excel_row + 1		'increment the excel row to add more
Next

'format the Worklist Excel
For col_to_autofit = 1 to emer_col
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

If developer_mode = False Then
	On Error Resume Next
	For cow = 0 to Ubound(CASES_TO_ASSIGN_ARRAY, 2)
		Err.Clear
		mvFSO.MoveFile CASES_TO_ASSIGN_ARRAY(file_path_const, cow), interview_team_cases_already_on_worklist & "\" & CASES_TO_ASSIGN_ARRAY(file_name_const, cow)

		If Err.Number = 70 Then
			attempt = 0
			Do
				EMWaitReady 0,0
				Err.Clear
				mvFSO.MoveFile CASES_TO_ASSIGN_ARRAY(file_path_const, cow), interview_team_cases_already_on_worklist & "\" & CASES_TO_ASSIGN_ARRAY(file_name_const, cow)
				attempt = attempt + 1
			Loop until Err.Number = 0 or attempt = 6
			If Err.Number <> 0 Then
				email_subject = "LD Interview Worklist MOVE FILE ERROR"
				email_body = "Go manually move file: " & CASES_TO_ASSIGN_ARRAY(file_path_const, cow)
				call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
			End If
		End If
	Next
	Err.Clear
	On Error Goto 0

	'Saves and closes the Excel Worklist - Naming convention is 'Interview Team Cases Worklist from MM-DD_MM-DD_MM-DD.xlsx' with all interview dates listed
	' formatted_date = replace(date, "/", "-")
	list_of_dates = " "
	For chkn_wg = 0 to UBound(DATES_WITH_INTERVIEWS_ARRAY, 2)
		If DATES_WITH_INTERVIEWS_ARRAY(checkbox_const, chkn_wg) = checked Then list_of_dates = list_of_dates & DATES_WITH_INTERVIEWS_ARRAY(month_const, chkn_wg) & "-" & DATES_WITH_INTERVIEWS_ARRAY(day_const, chkn_wg) & " "
	Next
	list_of_dates = replace(trim(list_of_dates), " ", "_")

	Set FSOxl = CreateObject("Scripting.FileSystemObject")
	base_file_name = worklist_folder & "\Interview Team Cases Worklist from " & list_of_dates
	worklist_file_name = worklist_folder & "\Interview Team Cases Worklist from " & list_of_dates
	full_worklist_file_name = worklist_file_name & ".xlsx"
	file_numb_count = 1
	Do
		file_is_already_here = False
		file_is_already_here = FSOxl.FileExists(full_worklist_file_name)
		If file_is_already_here Then
			worklist_file_name = base_file_name & "_" & file_numb_count
			full_worklist_file_name = worklist_file_name & ".xlsx"
			file_numb_count = file_numb_count + 1
		End If
	Loop until file_is_already_here = False

	objExcel.ActiveWorkbook.SaveAs full_worklist_file_name
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

	end_message = end_message & vbCr & "Excel created to save the information of the interviews created today." & vbCr
End If

internal_run_time = timer - start_time
internal_run_min = int(internal_run_time/60)
internal_run_sec = internal_run_time MOD 60

end_message = end_message & vbCr & "Manager Log and Eligibility Staff Excel files updated with cases:"
end_message = end_message & vbCr & "Total Cases added to logs: " & total_cases_assigned_count
end_message = end_message & vbCr & "FAMILIES Cases added to logs: " & fam_cases_assigned_count
end_message = end_message & vbCr & "ADULTS Cases added to logs: " & adul_cases_assigned_count
end_message = end_message & vbCr & "Files left open for review but updates are complete." & vbCr
end_message = end_message & vbCr & "Script run time: " & internal_run_min & " min, " & internal_run_sec & " sec."

Call script_end_procedure(end_message)
