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

'Dialog is just to alert user that the script starts with information gathering and takes time.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 191, 105, "Interview Team Cases Worklist"
  ButtonGroup ButtonPressed
    OkButton 125, 80, 50, 15
  Text 10, 10, 175, 20, "This script will create a list of cases with interviews completed by the Interview Team."
  Text 10, 35, 165, 20, "First the script needs to gather some information, which may take a minute or two. "
  Text 10, 65, 145, 10, "Please be patient, the script is running!"
  Text 10, 80, 105, 10, "Press OK to start info gather."
EndDialog

'no options on the dialog so no loop is necessary
dialog Dialog1
cancel_confirmation

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
Const dates_last_const	= 10

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

'Here is the worklist creation section
'these constants are to document the columns  - using a constant supports future changes
Const case_numb_col 	= 1
Const intvw_date_col 	= 2
Const exp_det_col 		= 3
Const cash_col 			= 4
Const cash_type_col		= 5
Const grh_col 			= 6
Const snap_col 			= 7
Const emer_col 			= 8

'Creating the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first row with header information
ObjExcel.Cells(1, case_numb_col).Value = "CASE NUMBER"
ObjExcel.Cells(1, intvw_date_col).Value = "INTERVIEW COMPLETED"
ObjExcel.Cells(1, exp_det_col).Value = "APPEARS EXPEDITED"
ObjExcel.Cells(1, cash_col).Value = "CASH"
ObjExcel.Cells(1, cash_type_col).Value = "CASH TYPE"
ObjExcel.Cells(1, grh_col).Value = "GRH"
ObjExcel.Cells(1, snap_col).Value = "SNAP"
ObjExcel.Cells(1, emer_col).Value = "EMER"
FOR i = 1 to emer_col							'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT
excel_row = 2

'creating some objects needed for XML handling
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set xml = CreateObject("Msxml2.DOMDocument")

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
			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream
			If .FileExists(xmlPath) = True then
				xmlDoc.Async = False

				' Load the XML file
				xmlDoc.load(xmlPath)

				'reads data about the case from the XML
				set node = xmlDoc.SelectSingleNode("//CaseNumber")
				MAXIS_case_number = node.text

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

				'Add the file information to the Excel document for the worklist
				ObjExcel.Cells(excel_row, case_numb_col).Value = MAXIS_case_number
				ObjExcel.Cells(excel_row, intvw_date_col).Value = script_date
				If exp_det = True Then ObjExcel.Cells(excel_row, exp_det_col).Value = "Yes"
				If req_cash = True Then
					ObjExcel.Cells(excel_row, cash_col).Value = "True"
					ObjExcel.Cells(excel_row, cash_type_col).Value = cash_type
				End If
				If req_grh = True Then ObjExcel.Cells(excel_row, grh_col).Value = "True"
				If req_snap = True Then ObjExcel.Cells(excel_row, snap_col).Value = "True"
				If req_emer = True Then ObjExcel.Cells(excel_row, emer_col).Value = "True"

				'THIS IS NOT WORKING BUT WILL RECORD THE DATE AND TIME THE CASE IS ADDED TO A WORKLIST IN THE XML
				' Set root = xmlDoc.SelectSingleNode"interview"
				' Set root = xmlDoc.documentElement
				' Set root = xmlDoc.SelectSingleNode("//inteview")
				' Set element = xmlDoc.createElement("AddedToWorklist")
				' xmlDoc.DocumentElement.appendChild element
				' Set info = xmlDoc.createTextNode(now)
				' element.appendChild info

				' xml.Save xmlPath
				' MsgBox "PAUSE"

				excel_row = excel_row + 1		'increment the excel row to add more

				'moving each file to the folder for cases already in a worklist
				.MoveFile xmlPath , interview_team_cases_already_on_worklist & "\" & file_name & ".xml"
			End If
		End With
	End If
Next
set xmlDoc = nothing

'format the Worklist Excel
For col_to_autofit = 1 to emer_col
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

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

end_msg = "Interviews from " & list_of_dates & " have been added to a worklist."
end_msg = end_msg & vbCr & "Worklist Name: " & replace(full_worklist_file_name, worklist_folder & "\", "")
end_msg = end_msg & vbCr & vbCr & "Worklist has been left open and can be found here:"
end_msg = end_msg & vbCr & full_worklist_file_name
Call script_end_procedure(end_msg)
