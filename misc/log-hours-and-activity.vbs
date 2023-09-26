'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - DASH.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = true		'this script does NOT open Global Variables. Setting the runLocally here.
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

'FUNCTIONS==================================================================================================================
function format_time_variable(time_variable, is_this_from_excel)
	If is_this_from_excel = True Then time_variable = time_variable * 24
	time_hour = Int(time_variable)
	time_minute = time_variable - time_hour
	' MsgBox time_mi6ute
	time_minute = time_minute * 60
	time_variable = TimeSerial(time_hour, time_minute, 0)
end function

'===========================================================================================================================

'DECLARATIONS===============================================================================================================
git_hub_issue_button = 1001
switch_activity_button = 1002
start_break_button = 1003
end_work_day_button  = 1004
request_time_off_button = 1005
insert_leave = 1006
approve_leave_button = 1007
i_approve_thee_button = 1008
request_additional_detail_button = 1009

task_category_list = "Select One..."
task_category_list = task_category_list+chr(9)+"Admin"
task_category_list = task_category_list+chr(9)+"Agency Leadership"
task_category_list = task_category_list+chr(9)+"Break"
task_category_list = task_category_list+chr(9)+"BZST Strategy and Planning"
task_category_list = task_category_list+chr(9)+"Consulting and Discovery"
task_category_list = task_category_list+chr(9)+"Data Analysis"
task_category_list = task_category_list+chr(9)+"Inter-Agency Collaboration"
task_category_list = task_category_list+chr(9)+"Planning"
task_category_list = task_category_list+chr(9)+"Process Analysis and Revision"
task_category_list = task_category_list+chr(9)+"Teaching"
task_category_list = task_category_list+chr(9)+"Team Building"
task_category_list = task_category_list+chr(9)+"Training"
task_category_list = task_category_list+chr(9)+"Travel"
task_category_list = task_category_list+chr(9)+"Script Projects"
task_category_list = task_category_list+chr(9)+"Supervisory"


'===========================================================================================================================
'setting the current time and fifteen minutes from now for display/defaulting
now_time = time
end_time_hr = DatePart("h", now_time)
end_time_min = DatePart("n", now_time)
If end_time_min = 1 OR end_time_min = 2 Then end_time_min = 0
If end_time_min = 3 OR end_time_min = 4 OR end_time_min = 6 OR end_time_min = 7 Then end_time_min = 5
If end_time_min = 8 OR end_time_min = 9 OR end_time_min = 11 OR end_time_min = 12 Then end_time_min = 10
If end_time_min = 13 OR end_time_min = 14 OR end_time_min = 16 OR end_time_min = 17 Then end_time_min = 15
If end_time_min = 18 OR end_time_min = 19 OR end_time_min = 21 OR end_time_min = 22 Then end_time_min = 20
If end_time_min = 23 OR end_time_min = 24 OR end_time_min = 26 OR end_time_min = 27 Then end_time_min = 25
If end_time_min = 28 OR end_time_min = 29 OR end_time_min = 31 OR end_time_min = 32 Then end_time_min = 30
If end_time_min = 33 OR end_time_min = 34 OR end_time_min = 36 OR end_time_min = 37 Then end_time_min = 35
If end_time_min = 38 OR end_time_min = 39 OR end_time_min = 41 OR end_time_min = 42 Then end_time_min = 40
If end_time_min = 43 OR end_time_min = 44 OR end_time_min = 46 OR end_time_min = 47 Then end_time_min = 45
If end_time_min = 48 OR end_time_min = 49 OR end_time_min = 51 OR end_time_min = 52 Then end_time_min = 50
If end_time_min = 53 OR end_time_min = 54 OR end_time_min = 56 OR end_time_min = 57 Then end_time_min = 55
If end_time_min = 58 OR end_time_min = 59 Then
	end_time_min = 0
	end_time_hr = end_time_hr + 1
End If

now_time = TimeSerial(end_time_hr, end_time_min, 0)
fifteen_minutes_from_now = DateAdd("n", 15, now_time)
end_time = now_time & ""
the_year = DatePart("yyyy", date)

'Defining the excel files for when running the script
excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking"

If user_ID_for_validation = "ILFE001" Then
	t_drive_excel_file_path = excel_file_path & "\Ilse Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Ilse Time Tracking.xlsx"
	bz_member = "Ilse Ferris"
	' leave_request_type = "PTO"
End If
If user_ID_for_validation = "MEGE001" Then
	t_drive_excel_file_path = excel_file_path & "\Megan Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Megan Time Tracking.xlsx"
	bz_member = "Megan Geissler"
	' leave_request_type = "PTO"
End If
If user_ID_for_validation = "CALO001" Then
	t_drive_excel_file_path = excel_file_path & "\Casey Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Casey Time Tracking.xlsx"
	bz_member = "Casey Love"
	' leave_request_type = "Vacation"
End If
If user_ID_for_validation = "MARI001" Then
	t_drive_excel_file_path = excel_file_path & "\Mark Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Mark Time Tracking.xlsx"
	bz_member = "Mark Riegel"
	' leave_request_type = "PTO"
End If
If user_ID_for_validation = "DACO003" Then
	t_drive_excel_file_path = excel_file_path & "\Dave Time Tracking.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Dave Time Tracking.xlsx"
	bz_member = "Dave Courtright"
	' leave_request_type = "PTO"
End If


If my_docs_excel_file_path = "" Then Call script_end_procedure("We have not set up your Time Tracking Worksheet yet!")
If objFSO.FileExists(my_docs_excel_file_path) = False Then Call script_end_procedure("We have not set up your Time Tracking Worksheet yet!")

view_excel = True		'this variable allows us to set
Call excel_open(my_docs_excel_file_path, view_excel, False, ObjExcel, objWorkbook)		'opening the excel file
ObjExcel.worksheets("Active Time Tracking").Activate

' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
' time_off_excel_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking\BZST Time Off.xlsx"
' time_off_requestes_for_approval = False
' ' If user_ID_for_validation = "ILFE001" Then
' 	view_excel = False		'this variable allows us to set
' 	Call excel_open(time_off_excel_path, view_excel, False, ObjTimeOffExcel, objTimeOffWorkbook)		'opening the excel file
' 	excel_row = 2
' 	Do
' 		If ObjTimeOffExcel.Cells(excel_row, 1).Value <> "" Then
' 			Call read_boolean_from_excel(ObjTimeOffExcel.Cells(excel_row, 9).Value, case_approved_variable)
' 			If case_approved_variable = False Then
' 				time_off_requestes_for_approval = True
' 				Exit Do
' 			End If
' 		End If
' 		excel_row = excel_row + 1
' 	Loop until ObjTimeOffExcel.Cells(excel_row, 1).Value = ""
' 	ObjTimeOffExcel.Quit
' ' End If

on_task = False						'default the booleans
current_task_from_today = False
project_droplist = ""
'find the open ended line
excel_row = 2
Do
	row_date = ObjExcel.Cells(excel_row, 1).Value								'removing this from here for now so we read it after we find the last row
	' row_start_time = ObjExcel.Cells(excel_row, 2).Value
	' row_end_time = ObjExcel.Cells(excel_row, 3).Value
	' row_start_time = row_start_time * 24
	' row_end_time = row_end_time * 24
	' MsgBox row_date
	row_date = DateAdd("d", 0, row_date)
	If DateDiff("d", row_date, date) < 31 Then
		listed_project = trim(ObjExcel.Cells(excel_row, 9).Value)
		If listed_project <> "" Then
			If InStr(project_droplist, listed_project) = 0 Then
				project_droplist = project_droplist & "~!~" & listed_project
			End If
		End If
	End If
	listed_project = ""
	' If row_date = date Then
	' 		time_spent = ObjExcel.Cells(excel_row, 4).Value
	' 		If time_spent <> "" Then
	' 		time_spent_hour = DatePart("h", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))					'here we create a number of the time spend so we can add it together
	' 		time_spent_min = DatePart("n", TIME_TRACKING_ARRAY(activity_time_spent, activity_count))
	' 		time_spent_min = time_spent_min/60
	' 		TIME_TRACKING_ARRAY(activity_time_spent_val, activity_count) = time_spent_hour + time_spent_min				'saving the time spent value into the array
	' 	End If
	' End If
	' If IsNumeric(row_start_time) = True and row_end_time = "" Then
	' 	on_task = True
	' 	If DateDiff("d", row_date, date) = 0 then current_task_from_today = True
	' End If
	excel_row = excel_row + 1
	next_row_date = ObjExcel.Cells(excel_row, 1).Value
Loop until next_row_date = ""
project_droplist = replace(project_droplist, "~!~", chr(9))

row_date = ObjExcel.Cells(excel_row - 1, 1).Value								'excel_row is now the first blank row, so we read the last row
row_start_time = ObjExcel.Cells(excel_row - 1, 2).Value
row_end_time = ObjExcel.Cells(excel_row - 1, 3).Value
row_date = DateAdd("d", 0, row_date)
If IsNumeric(row_start_time) = True and row_end_time = "" Then					'this is how we determine if we are on a task or not
	on_task = True
	If DateDiff("d", row_date, date) = 0 then current_task_from_today = True
End If
If on_task = True Then															'if we are on a task, this is going to read the information about the current task
	current_task_row = excel_row - 1
	current_category = ObjExcel.Cells(current_task_row, 5).Value
	current_meeting = ObjExcel.Cells(current_task_row, 6).Value
	current_detail = ObjExcel.Cells(current_task_row, 7).Value
	current_gh_issue = ObjExcel.Cells(current_task_row, 8).Value
	current_project = ObjExcel.Cells(current_task_row, 9).Value
End If
next_blank_row = excel_row			'setting the row number to know where to put new information

If row_start_time <> "" Then call format_time_variable(row_start_time, True)	'making these thing seasier to read and enter
If row_end_time <> "" Then call format_time_variable(row_end_time, True)

If current_meeting = "" Then current_meeting = "No"								'filling in a blank

current_elapsed_time = time - row_start_time									'calculating how much time has been spent on the furrent task
current_elapsed_time = current_elapsed_time * 24
elapsed_hr = Int(current_elapsed_time)
elapsed_min = current_elapsed_time - elapsed_hr
If len(elapsed_min) > 5 Then elapsed_min = left(elapsed_min, 5)
current_elapsed_time = elapsed_hr + elapsed_min
elapsed_min = elapsed_min * 60
elapsed_min = Int(elapsed_min)
elapsed_time_string = elapsed_hr & " hr, " & elapsed_min & " min"
' Call format_time_variable(current_elapsed_time, True)

' MsgBox "Start time - " & row_start_time & vbCr &_								'TESTING Code
'        "Elapsed_time - " & current_elapsed_time & vbCr &_
' 	   elapsed_time_string

' MsgBox "Date - " & row_date & vbCr & "Start time - " & row_start_time & vbCr &_
' 	   "End time - " & row_end_time & vbCr &_
' 	   "On Task - " & on_task & vbCr &_
' 	   "row - " & current_task_row & vbCr &_
' 	   "current_category - " & current_category & vbCr &_
' 	   "current_meeting - " & current_meeting & vbCr &_
' 	   "current_detail - " & current_detail & vbCr &_
' 	   "current_gh_issue - " & current_gh_issue & vbCr &_
' 	   "current_project - " & current_project & vbCr & vbCr &_
' 	   "on_task - " & on_task & vbCr &_
' 	   "current_task_from_today - " & current_task_from_today

objExcel.Visible = True															'showing the Excel File

If on_task = True and current_task_from_today = False Then						'this is if the open ended task is from a different day - we need to close that off first.
	end_date = row_date & ""			'defaulting the end date
	Do
		err_msg = ""
		BeginDialog Dialog1, 0, 0, 221, 180, "Work Day End"
		  EditBox 65, 140, 50, 15, end_date
		  EditBox 165, 140, 50, 15, end_time
		  ButtonGroup ButtonPressed
		    If current_gh_issue <> "" Then PushButton 20, 115, 115, 15, "GitHub Issue #" & current_gh_issue, git_hub_issue_button
		    OkButton 115, 160, 50, 15
		    CancelButton 165, 160, 50, 15
		  Text 10, 10, 195, 10, "It looks as though you didn't end your work day yesterday."
		  Text 10, 25, 105, 10, "When did you finish this task:"
		  Text 20, 40, 195, 10, "Category: " & current_category
		  Text 20, 55, 195, 10, "Detail: " & current_detail
		  Text 20, 70, 75, 10, "Meeting: " & current_meeting
		  Text 20, 85, 170, 10, "Project: " & current_project
		  Text 20, 100, 85, 10, "Start Time: " & row_start_time
		  Text 20, 145, 40, 10, "End Date:"
		  Text 125, 145, 35, 10, "End Time:"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then ObjExcel.Quit
		cancel_without_confirmation

		'making sure we are putting in date information
		If IsDate(end_date) = False Then err_msg = err_msg & " - Enter the next date as a valid date."
		If IsDate(end_time) = False Then err_msg = err_msg & " - Enter the time as a valid time."

		If ButtonPressed = git_hub_issue_button Then
			run "C:\Program Files\Google\Chrome\Application\chrome.exe https://github.com/Hennepin-County/MAXIS-scripts/issues/" & current_gh_issue
			err_msg = "LOOP"
		Else
			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg
		End If
	Loop until err_msg = ""
	end_date = DateAdd("d", 0, end_date)			'making this a date instead of a string'

	If end_date <> row_date Then					'if we are not on the same day, the update needs to be manual because there are multiple lines to add
		objExcel.Visible = True
		end_msg = "You have indicated that your work ended a different day." & vbCr & vbCr & "This script cannot repair time tracking that are from other days. The script has made the excel file visible. Update it manually, tracking past work. Be sure to save the file." & vbCr & vbCr & "The script will now end."
		Call script_end_procedure(end_msg)
	Else											'entering the end time and the time spent calculation
		ObjExcel.Cells(current_task_row, 3).Value = end_time
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"
		on_task = False								'resetting the on-task variable so that it can log in a new task next
	End If
End If

If on_task = True and current_task_from_today = True Then						'If we are on a task from today, we get the options for starting a new task, taking a break or ending the day.'
	'Dialog with information and buttons only to select the next steps
	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 361, 150, "Log Activity"
		  GroupBox 10, 25, 345, 100, "Activity in Progress"
		  Text 25, 45, 85, 10, "Date: " & row_date
		  Text 25, 65, 85, 10, "Start Time: " & row_start_time
		  Text 25, 85, 190, 10, "Category: " & current_category
		  Text 25, 105, 185, 10, "Detail: " & current_detail
		  Text 230, 45, 65, 10, "Meeting? " & current_meeting
		  Text 230, 65, 115, 10, "Project: " & current_project
		  Text 230, 105, 95, 10, "Elapsed Time: " & elapsed_time_string
		  ButtonGroup ButtonPressed
		    If current_gh_issue <> "" Then PushButton 260, 80, 85, 15, "GitHub Issue #" & current_gh_issue, git_hub_issue_button
		    PushButton 135, 5, 65, 15, "Switch Activity", switch_activity_button
		    PushButton 205, 5, 60, 15, "Start Break", start_break_button
		    PushButton 270, 5, 85, 15, "End Work Day", end_work_day_button
			' PushButton 10, 130, 70, 15, "Request Time Off", request_time_off_button													' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
			' PushButton 85, 130, 60, 15, "Insert Leave", insert_leave
			' If time_off_requestes_for_approval = True Then PushButton 150, 130, 60, 15, "Approve Leave", approve_leave_button			' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
		    OkButton 255, 130, 50, 15
		    CancelButton 305, 130, 50, 15
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then ObjExcel.Quit
		cancel_without_confirmation

		If ButtonPressed = git_hub_issue_button Then
			run "C:\Program Files\Google\Chrome\Application\chrome.exe https://github.com/Hennepin-County/MAXIS-scripts/issues/" & current_gh_issue
			err_msg = "LOOP"
		End If

		' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
		' If ButtonPressed = approve_leave_button Then
		' 	err_msg = "LOOP"
		' 	view_excel = True 		'this variable allows us to set
		' 	Call excel_open(time_off_excel_path, view_excel, False, ObjTimeOffExcel, objTimeOffWorkbook)		'opening the excel file
		' 	excel_row = 2
		' 	Do
		' 		If ObjTimeOffExcel.Cells(excel_row, 1).Value <> "" Then
		' 			Call read_boolean_from_excel(ObjTimeOffExcel.Cells(excel_row, 9).Value, case_approved_variable)
		' 			If case_approved_variable = False Then
		'
		' 				bz_member = ObjTimeOffExcel.Cells(excel_row, 1).Value
		' 				leave_request_start_date = ObjTimeOffExcel.Cells(excel_row, 2).Value
		' 				leave_request_start_time = ObjTimeOffExcel.Cells(excel_row, 3).Value & ""
		' 				leave_request_end_date = ObjTimeOffExcel.Cells(excel_row, 4).Value
		' 				leave_request_end_time = ObjTimeOffExcel.Cells(excel_row, 5).Value & ""
		' 				leave_request_days_off = ObjTimeOffExcel.Cells(excel_row, 6).Value & ""
		' 				' ObjTimeOffExcel.Cells(excel_row, 7).Value =
		' 				leave_request_type = ObjTimeOffExcel.Cells(excel_row, 8).Value
		' 				leave_request_notes = ObjTimeOffExcel.Cells(excel_row, 10).Value & ""
		'
		' 				Dialog1 = ""
		' 				BeginDialog Dialog1, 0, 0, 381, 170, "Time Off Request"
		' 				  EditBox 5, 125, 370, 15, email_comments
		' 				  ButtonGroup ButtonPressed
		' 				    PushButton 305, 150, 70, 15, "I Approve Thee", i_approve_thee_button
		' 				    PushButton 5, 150, 95, 15, "Request More Info", request_additional_detail_button
		' 				  GroupBox 5, 5, 370, 105, "Leave Request Information"
		' 				  Text 15, 20, 60, 10, "Request Person:"
		' 				  Text 75, 20, 60, 10, bz_member
		' 				  Text 30, 30, 40, 10, "Leave Type:"
		' 				  Text 75, 30, 60, 10, leave_request_type
		' 				  Text 155, 30, 75, 10, "Total WORK Days off:"
		' 				  Text 235, 30, 25, 10, leave_request_days_off
		' 				  Text 15, 45, 40, 10, "Start Date:"
		' 				  Text 55, 45, 50, 10, leave_request_start_date
		' 				  Text 25, 55, 40, 10, "Start Time:"
		' 				  Text 65, 55, 50, 10, leave_request_start_time
		' 				  Text 145, 45, 35, 10, "End Date:"
		' 				  Text 180, 45, 50, 10, leave_request_end_date
		' 				  Text 155, 55, 35, 10, "End Time:"
		' 				  Text 195, 55, 50, 10, leave_request_end_time
		' 				  Text 15, 70, 30, 10, "Notes:"
		' 				  Text 15, 80, 355, 25, leave_request_notes
		' 				  Text 5, 115, 80, 10, "Comments for the Email"
		' 				EndDialog
		'
		' 				dialog Dialog1
		'
		' 				If ButtonPressed = i_approve_thee_button Then
		'
		' 					'Creating a document of the request
		' 					Set objWord = CreateObject("Word.Application")
		'
		' 					objWord.Caption = "Leaave Request - " & bz_member & " - " & leave_request_start_date & " thru " & leave_request_end_date
		'
		' 					objWord.Visible = True														'Let the worker see the document
		'
		' 					Set objDoc = objWord.Documents.Add()										'Start a new document
		' 					Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document
		'
		' 					objSelection.Font.Name = "Arial"											'Setting the font before typing
		' 					objSelection.Font.Size = "20"
		' 					objSelection.Font.Bold = TRUE
		' 					objSelection.TypeText "Leave Request"
		' 					objSelection.TypeParagraph()
		' 					objSelection.Font.Size = "14"
		' 					objSelection.Font.Bold = FALSE
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "Check One:"
		' 					objSelection.Font.ColorIndex = 1
		' 					objSelection.TypeParagraph()
		'
		' 					If leave_request_type = "Sick" Then objSelection.TypeText ChrW(9746)
		' 					If leave_request_type <> "Sick" Then objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText " Sick"
		' 					objSelection.TypeText Chr(9)
		'
		' 					If leave_request_type = "Vacation" Then objSelection.TypeText ChrW(9746)
		' 					If leave_request_type <> "Vacation" Then objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText " Vacation"
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.TypeText Chr(9)
		'
		' 					objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText " * Conference"
		' 					objSelection.TypeText Chr(9)
		'
		' 					If leave_request_type = "Unpaid" Then objSelection.TypeText ChrW(9746)
		' 					If leave_request_type <> "Unpaid" Then objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText " SLWOP"
		' 					objSelection.TypeText Chr(9)
		'
		' 					If leave_request_type = "PTO" OR leave_request_type = "FMLA"  OR leave_request_type = "Holiday" Then objSelection.TypeText ChrW(9746)
		' 					If leave_request_type <> "PTO" AND leave_request_type <> "FMLA"  AND leave_request_type <> "Holiday" Then objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText " * Other"
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "1 day or less - Date: "
		' 					objSelection.Font.ColorIndex = 0
		' 					If DateDiff("d", leave_request_start_date, leave_request_end_date) = 0 Then
		' 						leave_request_start_date = leave_request_start_date & ""
		' 						objSelection.TypeText leave_request_start_date
		' 					Else
		' 						objSelection.TypeText Chr(9)
		' 						objSelection.TypeText Chr(9)
		' 					End If
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "Time: From "
		' 					objSelection.Font.ColorIndex = 0
		' 					' MsgBox leave_request_start_time
		' 					objSelection.TypeText leave_request_start_time
		' 					If leave_request_start_time = "" Then
		' 						objSelection.TypeText Chr(9)
		' 						objSelection.TypeText Chr(9)
		' 					End If
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "to "
		' 					objSelection.Font.ColorIndex = 0
		' 					objSelection.TypeText leave_request_end_time
		' 					If leave_request_end_time = "" Then
		' 						objSelection.TypeText Chr(9)
		' 						objSelection.TypeText Chr(9)
		' 					End If
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "If more than 1 Day - Date: "
		' 					objSelection.Font.ColorIndex = 0
		' 					If DateDiff("d", leave_request_start_date, leave_request_end_date) <> 0 Then
		' 						leave_request_start_date = leave_request_start_date & ""
		' 						leave_request_end_date = leave_request_end_date & ""
		' 						objSelection.TypeText leave_request_start_date
		' 						objSelection.TypeText Chr(9)
		' 						objSelection.Font.ColorIndex = 2
		' 						objSelection.TypeText "Through "
		' 						objSelection.Font.ColorIndex = 0
		' 						objSelection.TypeText leave_request_end_date
		' 					End If
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "Total number of work days off ( "
		' 					objSelection.Font.ColorIndex = 0
		' 					objSelection.TypeText leave_request_days_off
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText " ) "
		' 					objSelection.Font.ColorIndex = 0
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "* Explain Conference or Other"
		' 					objSelection.Font.ColorIndex = 0
		' 					objSelection.TypeParagraph()
		' 					objSelection.TypeParagraph()
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.Bold = TRUE
		' 					objSelection.Font.Underline = 1
		' 					objSelection.TypeText "SUPERVISOR WILL RESPOND BY EDITING SENDER'S E-MAIL"
		' 					objSelection.Font.Underline = 0
		' 					objSelection.TypeParagraph()
		' 					objSelection.TypeText ChrW(9746)
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.TypeText "Approved (Approval is conditional upon having sufficient leave hours at the time the leave is taken.)"
		' 					objSelection.TypeParagraph()
		' 					objSelection.TypeText ChrW(9744)
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.TypeText "Denied"
		' 					objSelection.Font.Bold = False
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.Font.ColorIndex = 2
		' 					objSelection.TypeText "Reason for Denial:"
		' 					objSelection.TypeParagraph()
		' 					objSelection.TypeParagraph()
		'
		' 					objSelection.TypeText "This request fits the Family Medical Leave Act eligibility criteria."
		' 					objSelection.TypeParagraph()
		' 					objSelection.Font.ColorIndex = 0
		' 					objSelection.Font.Bold = TRUE
		' 					If leave_request_type = "FMLA" Then objSelection.TypeText ChrW(9673)
		' 					If leave_request_type <> "FMLA" Then objSelection.TypeText ChrW(9678)
		' 					objSelection.TypeText " Yes"
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.TypeText Chr(9)
		' 					objSelection.TypeText ChrW(9678)
		' 					objSelection.TypeText " No"
		' 					objSelection.Font.Bold = False
		'
		' 					objSelection.TypeParagraph()
		'
		' 					MsgBox "PAUSE 1"
		'
		'
		' 					'We set the file path and name based on case number and date. We can add other criteria if important.
		' 					'This MUST have the 'pdf' file extension to work
		' 					start_date_for_doc_file = replace(leave_request_start_date, "/", "-")
		' 					end_date_forr_doc_file = replace(leave_request_end_date, "/", "-")
		' 					pdf_doc_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking\Leave Time Documents\" & "Leaave Request - " & bz_member & " - " & start_date_for_doc_file & " thru " & end_date_forr_doc_file & ".pdf"
		'
		' 					'Now we save the document.
		' 					'MS Word allows us to save directly as a PDF instead of a DOC.
		' 					'the file path must be PDF
		' 					'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
		' 					objDoc.SaveAs pdf_doc_path, 17
		'
		' 					If objFSO.FileExists(pdf_doc_path) = TRUE Then
		' 						'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
		' 						objDoc.Close wdDoNotSaveChanges
		' 						objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)
		' 					End If
		'
		' 					ObjTimeOffExcel.Cells(excel_row, 9).Value = True
		'
		' 					objTimeOffWorkbook.Save
		'
		' 					' Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
		'
		' 					time_specific = True
		' 					If leave_request_start_time = "" OR leave_request_end_time = "" Then time_specific = False
		'
		' 					'Assigning needed numbers as variables for readability
		' 					olAppointmentItem = 1
		' 					olRecursDaily = 0
		'
		' 					'Creating an Outlook object item
		' 					Set objOutlook = CreateObject("Outlook.Application")
		' 					Set objAppointment = objOutlook.CreateItem(olAppointmentItem)
		'
		' 					'Assigning individual appointment options
		' 					If time_specific = True Then
		' 						objAppointment.Start = leave_request_start_date & " " & leave_request_start_time	'Start date and time are carried over from parameters
		' 						objAppointment.End = leave_request_end_date & " " & leave_request_end_time			'End date and time are carried over from parameters
		' 						objAppointment.AllDayEvent = False 													'Defaulting to false for this. Perhaps someday this can be true. Who knows.
		' 					Else
		' 						objAppointment.Start = leave_request_start_date & " 7:00"
		' 						objAppointment.End = leave_request_end_date & " 17:00"
		' 						objAppointment.AllDayEvent = True 													'Defaulting to false for this. Perhaps someday this can be true. Who knows.
		' 					End If
		'
		' 					objAppointment.BusyStatus = 0 'Free
		'
		' 					' olBusy	2	The user is busy.
		' 					' olFree	0	The user is available.
		' 					' olOutOfOffice	3	The user is out of office.
		' 					' olTentative	1	The user has a tentative appointment scheduled.
		' 					' olWorkingElsewhere	4	The user is working in a location away from the office.
		'
		' 					objAppointment.Sensitivity = 0	'Normal
		'
		' 					' olConfidential	3	Confidential
		' 					' olNormal	0	Normal sensitivity
		' 					' olPersonal	1	Personal
		' 					' olPrivate	2	Private
		'
		' 					objAppointment.Categories = "Staff Leave"
		'
		' 					If time_specific = True Then objAppointment.Subject = bz_member & " Out"
		' 					If time_specific = False Then objAppointment.Subject = bz_member & " OFD"							'Defining the subject from parameters
		' 					' objAppointment.Body = appt_body									'Defining the body from parameters
		' 					' objAppointment.Location = appt_location							'Defining the location from parameters
		' 					objAppointment.ReminderSet = False
		'
		' 					' objAppointment.Categories = appt_category						'Defines a category
		' 					objAppointment.Save												'Saves the appointment
		'
		' 					MsgBox "PAUSE 2"
		' 				End If
		'
		' 				If ButtonPressed = request_additional_detail_button Then
		'
		' 				End If
		'
		' 			End If
		' 		End If
		' 		excel_row = excel_row + 1
		' 	Loop until ObjTimeOffExcel.Cells(excel_row, 1).Value = ""
		' 	ObjTimeOffExcel.Quit
		' End If
	Loop until err_msg = ""
End If



	'Display the Leave information

	'Document the approval
	'Update the Excel
	'Create a PDF of the leave

	'Create a calendar obect

' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
' If ButtonPressed = request_time_off_button Then
' 	ObjExcel.Quit
' 	Do
' 		err_msg = ""
' 		BeginDialog Dialog1, 0, 0, 256, 165, "Time Off Request"
' 		  EditBox 90, 20, 50, 15, leave_request_start_date
' 		  EditBox 195, 20, 50, 15, leave_request_start_time
' 		  EditBox 90, 55, 50, 15, leave_request_end_date
' 		  EditBox 195, 55, 50, 15, leave_request_end_time
' 		  DropListBox 60, 90, 60, 45, "Vacation"+chr(9)+"PTO"+chr(9)+"Holiday"+chr(9)+"Sick"+chr(9)+"FMLA"+chr(9)+"Unpaid", leave_request_type
' 		  EditBox 220, 90, 25, 15, leave_request_days_off
' 		  EditBox 15, 120, 230, 15, leave_request_notes
' 		  ButtonGroup ButtonPressed
' 		    CancelButton 145, 140, 50, 15
' 		    OkButton 195, 140, 50, 15
' 		  GroupBox 5, 5, 245, 155, "Enter details about the time off you are requesting"
' 		  Text 50, 25, 40, 10, "Start Date:"
' 		  Text 155, 25, 40, 10, "Start Time:"
' 		  Text 120, 40, 130, 10, "(Date is required but time can be blank)"
' 		  Text 50, 60, 40, 10, "End Date:"
' 		  Text 155, 60, 40, 10, "End Time:"
' 		  Text 120, 75, 130, 10, "(Date is required but time can be blank)"
' 		  Text 15, 95, 40, 10, "Leave Type:"
' 		  Text 145, 95, 75, 10, "Total WORK Days off:"
' 		  Text 15, 110, 30, 10, "Notes:"
' 		EndDialog
'
' 		dialog Dialog1
' 		cancel_without_confirmation
'
' 	Loop until err_msg = ""
'
' 	view_excel = False		'this variable allows us to set
' 	Call excel_open(time_off_excel_path, view_excel, False, ObjTimeOffExcel, objTimeOffWorkbook)		'opening the excel file
'
' 	excel_row = 2
' 	Do
' 		If ObjTimeOffExcel.Cells(excel_row, 1).Value = "" Then
' 			ObjTimeOffExcel.Cells(excel_row, 1).Value = bz_member
' 			ObjTimeOffExcel.Cells(excel_row, 2).Value = leave_request_start_date
' 			ObjTimeOffExcel.Cells(excel_row, 3).Value = leave_request_start_time
' 			ObjTimeOffExcel.Cells(excel_row, 4).Value = leave_request_end_date
' 			ObjTimeOffExcel.Cells(excel_row, 5).Value = leave_request_end_time
' 			ObjTimeOffExcel.Cells(excel_row, 6).Value = leave_request_days_off
' 			' ObjTimeOffExcel.Cells(excel_row, 7).Value =
' 			ObjTimeOffExcel.Cells(excel_row, 8).Value = leave_request_type
' 			ObjTimeOffExcel.Cells(excel_row, 9).Value = False
' 			ObjTimeOffExcel.Cells(excel_row, 10).Value = leave_request_notes
' 			Exit Do
' 		End If
' 		excel_row = excel_row + 1
' 	Loop
' 	objTimeOffWorkbook.Save
' 	ObjTimeOffExcel.Quit
' 	end_msg = "Your request for time off has been added to the list for approval:" & vbCr & vbCr & "Start Date: " & leave_request_start_date & vbCr &_
' 			  "Start Time: " & leave_request_start_time & vbCr &_
' 			  "End Date: " & leave_request_end_date & vbCr &_
' 			  "End Time: " & leave_request_end_time & vbCr & vbCr &_
' 			  "Number of days off: " & leave_request_days_off
' 	Call script_end_procedure(end_msg)
' End If

If on_task = False Then					'If we are not currently on a task, this will start a new activity only
	next_date = date & ""				'defaulting the end time and date
	next_start_time = now_time & ""
	Do
		err_msg = ""
		BeginDialog Dialog1, 0, 0, 361, 135, "Log Activity"
		  GroupBox 10, 10, 345, 100, "Log New Activity"
		  EditBox 50, 25, 50, 15, next_date
		  EditBox 65, 45, 50, 15, next_start_time
		  DropListBox 65, 65, 155, 45, task_category_list, next_category
		  EditBox 50, 85, 170, 15, next_detail
		  DropListBox 265, 25, 30, 45, "?"+chr(9)+"Yes", next_meeting
		  ComboBox 260, 45, 90, 15, project_droplist, next_project
		  EditBox 280, 65, 35, 15, next_gh_issue
		  ButtonGroup ButtonPressed
			PushButton 10, 115, 60, 15, "Insert Leave", insert_leave
			OkButton 255, 115, 50, 15
			CancelButton 305, 115, 50, 15
		  Text 25, 30, 20, 10, "Date: "
		  Text 25, 50, 40, 10, "Start Time:"
		  Text 25, 70, 35, 10, "Category: "
		  Text 25, 90, 25, 10, "Detail:"
		  Text 230, 30, 30, 10, "Meeting"
		  Text 230, 50, 30, 10, "Project:"
		  Text 230, 70, 45, 10, "GitHub Issue:"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then ObjExcel.Quit
		cancel_without_confirmation

		If IsDate(next_date) = False Then err_msg = err_msg & " - Enter the next date as a valid date."
		If IsDate(next_start_time) = False Then err_msg = err_msg & " - Enter the time as a valid time."
		If next_category = "Select One..." Then err_msg = err_msg & " - Select the category"
		If next_gh_issue <> "" and IsNumeric(next_gh_issue) = False Then err_msg = err_msg & " - Enter the GitHub issue as the number only."

		If ButtonPressed = insert_leave Then err_msg = ""
		If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg

	Loop until err_msg = ""
	If next_meeting = "?" Then next_meeting = ""

	If ButtonPressed <> insert_leave Then
		'adding in the information from the new activity
		ObjExcel.Cells(next_blank_row, 1).Value = next_date
		If next_start_time <> now_time Then ObjExcel.Cells(next_blank_row, 2).Value = next_start_time
		ObjExcel.Cells(next_blank_row, 4).Value = ""								'we will be blanking this out because it will default to a formula that could cause errors on future runs of the script
		ObjExcel.Cells(next_blank_row, 5).Value = next_category
		ObjExcel.Cells(next_blank_row, 6).Value = next_meeting
		ObjExcel.Cells(next_blank_row, 7).Value = next_detail
		ObjExcel.Cells(next_blank_row, 8).Value = ""
		If next_gh_issue <> "" Then ObjExcel.Cells(next_blank_row, 8).Value = "=HYPERLINK(" & chr(34) & "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & next_gh_issue & chr(34) & ", " & chr(34) & next_gh_issue & chr(34) & ")"
		ObjExcel.Cells(next_blank_row, 9).Value = next_project
		ObjExcel.Cells(next_blank_row, 10).Value = "Y"
		end_msg = "As of " & next_date & " at " & next_start_time & " you are now working on:" & vbCr & "  - Category: " & next_category & vbCr & "  - Detail: " & next_detail
	End If
End If

' 'COMMENTING THIS OUT AS WE ARE NOT USING IT. NOT DELETING AS WE MAY DECIDE TO BRING IT BACK.
' If ButtonPressed = insert_leave Then
'
' End If

'If the on task option selected was to take a new action - this will start that
If ButtonPressed = switch_activity_button or ButtonPressed = start_break_button or ButtonPressed = end_work_day_button Then
	If ButtonPressed = end_work_day_button Then									'Ending the work day
		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 331, 45, "Work Day End"
			  EditBox 275, 5, 50, 15, end_time
			  ButtonGroup ButtonPressed
			    OkButton 225, 25, 50, 15
			    CancelButton 275, 25, 50, 15
			  Text 10, 10, 215, 10, "End of the work day! When have you finished?"
			  Text 235, 10, 35, 10, "End Time:"
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then ObjExcel.Quit
			cancel_without_confirmation

			If IsDate(end_time) = False Then err_msg = err_msg & vbCr & " - Enter the time as a valid time."

			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg
		Loop until err_msg = ""

		ObjExcel.Cells(current_task_row, 3).Value = end_time					'enter the end time and calculation
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		total_time_spent_today = 0
		excel_row = current_task_row
		Do
			row_date = ObjExcel.Cells(excel_row, 1).Value
			If row_date = date Then
				time_spent = ObjExcel.Cells(excel_row, 4).Value
				paid_yn = ObjExcel.Cells(excel_row, 10).Value

				If time_spent <> "" and paid_yn = "Y" Then
					time_spent_hour = DatePart("h", time_spent)					'here we create a number of the time spend so we can add it together
					time_spent_min = DatePart("n", time_spent)
					time_spent_min = time_spent_min/60
					time_spent_val = time_spent_hour + time_spent_min				'saving the time spent value into the array
					total_time_spent_today = total_time_spent_today + time_spent_val
				End If
			End If
			excel_row = excel_row - 1
		Loop Until row_date <> date

		end_msg = "Your work day has ended at " & end_time & vbCr & vbCr & "You have worked a total of " & total_time_spent_today & " hours today." & vbCr & vbCr & "Have a wonderful rest of the day!"
	End If

	If ButtonPressed = start_break_button Then									'starting a break
		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 331, 95, "Break Time"
			  EditBox 45, 35, 50, 15, end_time
			  DropListBox 265, 35, 60, 45, "Yes"+chr(9)+"No", break_yn
			  EditBox 45, 55, 280, 15, next_detail
			  ButtonGroup ButtonPressed
			    OkButton 225, 75, 50, 15
			    CancelButton 275, 75, 50, 15
			  Text 10, 10, 315, 20, "Current Task: " & current_category & " : " & current_detail
			  Text 10, 40, 35, 10, "End Time:"
			  Text 175, 40, 90, 10, "Is this break a paid break?"
			  Text 20, 60, 25, 10, "Detail:"
			EndDialog


			dialog Dialog1
			If ButtonPressed = 0 Then ObjExcel.Quit
			cancel_without_confirmation

			If IsDate(end_time) = False Then err_msg = err_msg & vbCr & " - Enter the time as a valid time."

			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg
		Loop until err_msg = ""

		ObjExcel.Cells(current_task_row, 3).Value = end_time					'enter the end time and time spent calculation for the current activity
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		If break_yn = "Yes" Then												'this is the Yes/No option of if the break is a paid break or not.
			ObjExcel.Cells(next_blank_row, 10).Value = "Y"						'entering the paid code
			If trim(next_detail) = "" Then										'if no detail is entered - it will enter 'Paid'
				ObjExcel.Cells(next_blank_row, 7).Value = "Paid"
			Else
				ObjExcel.Cells(next_blank_row, 7).Value = next_detail			'enter the detail that was entered into the dialog
			End If
			fifteen_minutes_from_break_start = DateAdd("n", 15, end_time)
			end_msg = "Your activity (Category: " & current_category & ", Details: " & current_detail & ") has been ended as of " & end_time & " and you are now on paid break." & vbCr & vbCR & "Fifteen minutes will be up at " & fifteen_minutes_from_break_start
		Else
			ObjExcel.Cells(next_blank_row, 10).Value = "N"						'entering an 'N' for the paid code
			If trim(next_detail) = "" Then
				ObjExcel.Cells(next_blank_row, 7).Value = "NOT PAID"			'if no detail is entered - it will enter 'NOT PAID'
			Else
				ObjExcel.Cells(next_blank_row, 7).Value = next_detail			'entering the detail that was entered into the dialog
			End If
			end_msg = "Your activity (Category: " & current_category & ", Details: " & current_detail & ") has been ended as of " & end_time & " and you are now on break. This is indicated as NOT paid."
		End If
		ObjExcel.Cells(next_blank_row, 1).Value = date							'entering the date and 'Break' category in the next line
		ObjExcel.Cells(next_blank_row, 5).Value = "Break"						'the start time does not need to be entered because it fills in from the end time on the previous line
		ObjExcel.Cells(next_blank_row, 8).Value = ""							'blanking out of the GH Issue Line'
	End If

	If ButtonPressed = switch_activity_button Then								'ending one activity and starting another.'
		next_date = date & ""				'defaulting to the current time and date'
		next_start_time = now_time & ""

		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 361, 130, "Log Activity"
			  Text 10, 10, 215, 10, "Current Task: " & current_category
			  EditBox 275, 5, 50, 15, end_time
			  EditBox 50, 40, 50, 15, next_date
			  ' EditBox 65, 60, 50, 15, next_start_time
			  DropListBox 65, 60, 155, 45, task_category_list, next_category
			  EditBox 50, 80, 170, 15, next_detail
			  DropListBox 265, 40, 30, 45, "?"+chr(9)+"Yes", next_meeting
			  ComboBox 260, 60, 90, 15, project_droplist, next_project
			  EditBox 280, 80, 35, 15, next_gh_issue
			  ButtonGroup ButtonPressed
				OkButton 255, 110, 50, 15
				CancelButton 305, 110, 50, 15
			  Text 235, 10, 35, 10, "End Time:"
			  GroupBox 10, 25, 345, 80, "Log New Activity"
			  Text 25, 45, 20, 10, "Date: "
			  ' Text 25, 65, 40, 10, "Start Time:"
			  Text 25, 65, 35, 10, "Category: "
			  Text 25, 85, 25, 10, "Detail:"
			  Text 230, 45, 30, 10, "Meeting"
			  Text 230, 65, 30, 10, "Project:"
			  Text 230, 85, 45, 10, "GitHub Issue:"
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then ObjExcel.Quit
			cancel_without_confirmation

			If IsDate(next_date) = False Then err_msg = err_msg & vbCr & " - Enter the next date as a valid date."
			If IsDate(end_time) = False Then err_msg = err_msg & vbCr & " - Enter the time as a valid time."
			If next_category = "Select One..." Then err_msg = err_msg & vbCr & " - Select the category"
			If next_gh_issue <> "" and IsNumeric(next_gh_issue) = False Then err_msg = err_msg & vbCr & " - Enter the GitHub issue as the number only."

			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg

		Loop until err_msg = ""
		If next_meeting = "?" Then next_meeting = ""

		ObjExcel.Cells(current_task_row, 3).Value = end_time					'entering the end time and time spent calculation on the current activity
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		ObjExcel.Cells(next_blank_row, 1).Value = next_date						'entering the information for the new activity'
		ObjExcel.Cells(next_blank_row, 4).Value = ""
		ObjExcel.Cells(next_blank_row, 5).Value = next_category
		ObjExcel.Cells(next_blank_row, 6).Value = next_meeting
		ObjExcel.Cells(next_blank_row, 7).Value = next_detail
		ObjExcel.Cells(next_blank_row, 8).Value = ""
		If next_gh_issue <> "" Then ObjExcel.Cells(next_blank_row, 8).Value = "=HYPERLINK(" & chr(34) & "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & next_gh_issue & chr(34) & ", " & chr(34) & next_gh_issue & chr(34) & ")"
		ObjExcel.Cells(next_blank_row, 9).Value = next_project
		ObjExcel.Cells(next_blank_row, 10).Value = "Y"
		end_msg = "As of " & next_date & " at " & end_time & " you are now working on:" & vbCr & "  - Category: " & next_category & vbCr & "  - Detail: " & next_detail
	End If
End If

objWorkbook.Save									'saving the file to 'My Documents'
objWorkbook.SaveAs (t_drive_excel_file_path)		'saving the file to the T Drive
ObjExcel.Quit										'closing the Excel File'
call script_end_procedure(end_msg)
