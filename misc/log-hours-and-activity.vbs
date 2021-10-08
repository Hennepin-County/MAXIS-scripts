'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - DASH.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = true
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
function format_time_variable(time_variable, is_this_from_excel)
	If is_this_from_excel = True Then time_variable = time_variable * 24
	time_hour = Int(time_variable)
	time_minute = time_variable - time_hour
	' MsgBox time_mi6ute
	time_minute = time_minute * 60
	time_variable = TimeSerial(time_hour, time_minute, 0)
end function

git_hub_issue_button = 1001
switch_activity_button = 1002
start_break_button = 1003
end_work_day_button  = 1004

excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Time Tracking"

If user_ID_for_validation = "CALO001" Then
	t_drive_excel_file_path = excel_file_path & "\Casey Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Casey Time Tracking 2021.xlsx"
End If
If user_ID_for_validation = "ILFE001" Then
	t_drive_excel_file_path = excel_file_path & "\Ilse Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "Ilse Time Tracking 2021.xlsx"
End If
If user_ID_for_validation = "WFS395" Then
	t_drive_excel_file_path = excel_file_path & "\MiKayla Time Tracking 2021.xlsx"
	my_docs_excel_file_path = user_myDocs_folder & "MiKayla Time Tracking 2021.xlsx"
End If

Call excel_open(my_docs_excel_file_path, False, False, ObjExcel, objWorkbook)

on_task = False
current_task_from_today = False

'find the open ended line
excel_row = 2
Do
	row_date = ObjExcel.Cells(excel_row, 1).Value
	row_start_time = ObjExcel.Cells(excel_row, 2).Value
	row_end_time = ObjExcel.Cells(excel_row, 3).Value
	' row_start_time = row_start_time * 24
	' row_end_time = row_end_time * 24

	row_date = DateAdd("d", 0, row_date)
	If IsNumeric(row_start_time) = True and row_end_time = "" Then
		on_task = True
		If DateDiff("d", row_date, date) = 0 then current_task_from_today = True
	End If
	excel_row = excel_row + 1
	next_row_date = ObjExcel.Cells(excel_row, 1).Value
Loop until next_row_date = ""

If on_task = True Then
	current_task_row = excel_row - 1
	current_category = ObjExcel.Cells(current_task_row, 5).Value
	current_meeting = ObjExcel.Cells(current_task_row, 6).Value
	current_detail = ObjExcel.Cells(current_task_row, 7).Value
	current_gh_issue = ObjExcel.Cells(current_task_row, 8).Value
	current_project = ObjExcel.Cells(current_task_row, 9).Value
End If
next_blank_row = excel_row
' MsgBox row_start_time & vbCr & TimeValue(row_start_time)
' MsgBox Time
If row_start_time <> "" Then call format_time_variable(row_start_time, True)
If row_end_time <> "" Then call format_time_variable(row_end_time, True)

If current_meeting = "" Then current_meeting = "No"

current_elapsed_time = time - row_start_time
current_elapsed_time = current_elapsed_time * 24
elapsed_hr = Int(current_elapsed_time)
elapsed_min = current_elapsed_time - elapsed_hr
If len(elapsed_min) > 5 Then elapsed_min = left(elapsed_min, 5)
current_elapsed_time = elapsed_hr + elapsed_min
elapsed_min = elapsed_min * 60
elapsed_min = Int(elapsed_min)
elapsed_time_string = elapsed_hr & " hr, " & elapsed_min & " min"
' Call format_time_variable(current_elapsed_time, True)

' MsgBox "Start time - " & row_start_time & vbCr &_
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

objExcel.Visible = True

If on_task = True and current_task_from_today = False Then
	end_date = row_date & ""
	Do
		err_msg = ""
		BeginDialog Dialog1, 0, 0, 221, 180, "Work Day End"
		  EditBox 65, 140, 50, 15, end_date
		  EditBox 165, 140, 50, 15, end_time
		  ButtonGroup ButtonPressed
		    OkButton 115, 160, 50, 15
		    CancelButton 165, 160, 50, 15
		    If current_gh_issue <> "" Then PushButton 20, 115, 115, 15, "GitHub Issue #" & current_gh_issue, git_hub_issue_button
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


		If IsDate(end_date) = False Then err_msg = err_msg & " - Enter the next date as a valid date."
		If IsDate(end_time) = False Then err_msg = err_msg & " - Enter the time as a valid time."

		If ButtonPressed = git_hub_issue_button Then
			run "C:\Program Files\Google\Chrome\Application\chrome.exe https://github.com/Hennepin-County/MAXIS-scripts/issues/" & current_gh_issue
			err_msg = "LOOP"
		Else
			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg
		End If
	Loop until err_msg = ""
	end_date = DateAdd("d", 0, end_date)

	If end_date <> row_date Then
		objExcel.Visible = True
		end_msg = "You have indicated that your work ended a different day." & vbCr & vbCr & "This script cannot repair time tracking that are from other days. The script has made the excel file visible. Update it manually, tracking past work. Be sure to save the file." & vbCr & vbCr & "The script will now end."
		Call script_end_procedure(end_msg)
	Else
		ObjExcel.Cells(current_task_row, 3).Value = end_time
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"
		on_task = False
	End If
End If


If on_task = True and current_task_from_today = True Then

	Do
		err_msg = ""
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
	Loop until err_msg = ""
End If

now_time = time
end_time_hr = DatePart("h", now_time)
end_time_min = DatePart("n", now_time)
' end_time_min = end_time_min/60
now_time = TimeSerial(end_time_hr, end_time_min, 0)
fifteen_minutes_from_now = DateAdd("n", 15, now_time)
end_time = now_time & ""

If on_task = False Then
	next_date = date & ""
	next_start_time = now_time & ""
	Do
		err_msg = ""
		BeginDialog Dialog1, 0, 0, 361, 135, "Log Activity"
		  GroupBox 10, 10, 345, 100, "Log New Activity"
		  Text 25, 30, 20, 10, "Date: "
		  Text 25, 50, 40, 10, "Start Time:"
		  Text 25, 70, 35, 10, "Category: "
		  Text 25, 90, 25, 10, "Detail:"
		  Text 230, 30, 30, 10, "Meeting"
		  Text 230, 50, 30, 10, "Project:"
		  ButtonGroup ButtonPressed
			OkButton 255, 115, 50, 15
			CancelButton 305, 115, 50, 15
		  Text 230, 70, 45, 10, "GitHub Issue:"
		  EditBox 50, 25, 50, 15, next_date
		  EditBox 65, 45, 50, 15, next_start_time
		  DropListBox 65, 65, 155, 45, "Select One..."+chr(9)+"Admin"+chr(9)+"Break"+chr(9)+"Consulting on Systems and Processes"+chr(9)+"Department Wide Script Tools"+chr(9)+"New Projects and Script Development"+chr(9)+"Ongoing Script Support"+chr(9)+"Other"+chr(9)+"Personal Skills Development"+chr(9)+"Team Strategy Development"+chr(9)+"Training"+chr(9)+"Travel", next_category
		  EditBox 50, 85, 170, 15, next_detail
		  DropListBox 265, 25, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", next_meeting
		  EditBox 260, 45, 90, 15, next_project
		  EditBox 280, 65, 35, 15, next_gh_issue
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then ObjExcel.Quit
		cancel_without_confirmation

		If IsDate(next_date) = False Then err_msg = err_msg & " - Enter the next date as a valid date."
		If IsDate(next_start_time) = False Then err_msg = err_msg & " - Enter the time as a valid time."
		If next_category = "Select One..." Then err_msg = err_msg & " - Select the category"
		If next_meeting = "?" Then err_msg = err_msg & " - Indicate if the activity is a meeting."
		If next_gh_issue <> "" and IsNumeric(next_gh_issue) = False Then err_msg = err_msg & " - Enter the GitHub issue as the number only."

		If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg

	Loop until err_msg = ""
	ObjExcel.Cells(next_blank_row, 1).Value = next_date
	If next_start_time <> now_time Then ObjExcel.Cells(next_blank_row, 2).Value = next_start_time
	ObjExcel.Cells(next_blank_row, 4).Value = ""
	' ObjExcel.Cells(next_blank_row, 3).Value =
	ObjExcel.Cells(next_blank_row, 5).Value = next_category
	ObjExcel.Cells(next_blank_row, 6).Value = next_meeting
	ObjExcel.Cells(next_blank_row, 7).Value = next_detail
	ObjExcel.Cells(next_blank_row, 8).Value = "=HYPERLINK(" & chr(34) & "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & next_gh_issue & chr(34) & ", " & chr(34) & next_gh_issue & chr(34) & ")"
	ObjExcel.Cells(next_blank_row, 9).Value = next_project
	ObjExcel.Cells(next_blank_row, 10).Value = "Y"
	end_msg = "As of " & next_date & " at " & next_start_time & " you are now working on:" & vbCr & "  - Category: " & next_category & vbCr & "  - Detail: " & next_detail

End If

If ButtonPressed = switch_activity_button or ButtonPressed = start_break_button or ButtonPressed = end_work_day_button Then
	If ButtonPressed = end_work_day_button Then
		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 331, 45, "Work Day End"
			  ButtonGroup ButtonPressed
			    OkButton 225, 25, 50, 15
			    CancelButton 275, 25, 50, 15
			  Text 10, 10, 215, 10, "End of the work day! When have you finished?"
			  Text 235, 10, 35, 10, "End Time:"
			  EditBox 275, 5, 50, 15, end_time
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then ObjExcel.Quit
			cancel_without_confirmation

			If IsDate(end_time) = False Then err_msg = err_msg & vbCr & " - Enter the time as a valid time."

			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg
		Loop until err_msg = ""

		ObjExcel.Cells(current_task_row, 3).Value = end_time
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		end_msg = "Your work day has ended at " & end_time & vbCr & vbCr & "Have a wonderful rest of the day!"
	End If

	If ButtonPressed = start_break_button Then
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

		ObjExcel.Cells(current_task_row, 3).Value = end_time
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		If break_yn = "Yes" Then
			ObjExcel.Cells(next_blank_row, 10).Value = "Y"
			If trim(next_detail) = "" Then
				ObjExcel.Cells(next_blank_row, 7).Value = "Paid"
			Else
				ObjExcel.Cells(next_blank_row, 7).Value = next_detail
			End If
			end_msg = "Your activity (Category: " & current_category & ", Details: " & current_detail & ") has been ended as of " & end_time & " and you are now on paid break." & vbCr & vbCR & "Fifteen minutes will be up at " & fifteen_minutes_from_now
		Else
			ObjExcel.Cells(next_blank_row, 10).Value = "N"
			If trim(next_detail) = "" Then
				ObjExcel.Cells(next_blank_row, 7).Value = "NOT PAID"
			Else
				ObjExcel.Cells(next_blank_row, 7).Value = next_detail
			End If
			end_msg = "Your activity (Category: " & current_category & ", Details: " & current_detail & ") has been ended as of " & end_time & " and you are now on break. This is indicated as NOT paid."
		End If
		ObjExcel.Cells(next_blank_row, 1).Value = date
		ObjExcel.Cells(next_blank_row, 5).Value = "Break"
	End If


	If ButtonPressed = switch_activity_button Then
		next_date = date & ""
		next_start_time = now_time & ""

		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 361, 130, "Log Activity"
			  Text 10, 10, 215, 10, "Current Task: " & current_category
			  EditBox 275, 5, 50, 15, end_time
			  EditBox 50, 40, 50, 15, next_date
			  ' EditBox 65, 60, 50, 15, next_start_time
			  DropListBox 65, 60, 155, 45, "Select One..."+chr(9)+"Admin"+chr(9)+"Break"+chr(9)+"Consulting on Systems and Processes"+chr(9)+"Department Wide Script Tools"+chr(9)+"New Projects and Script Development"+chr(9)+"Ongoing Script Support"+chr(9)+"Other"+chr(9)+"Personal Skills Development"+chr(9)+"Supervisory"+chr(9)+"Team Strategy Development"+chr(9)+"Training"+chr(9)+"Travel", next_category
			  EditBox 50, 80, 170, 15, next_detail
			  DropListBox 265, 40, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", next_meeting
			  EditBox 260, 60, 90, 15, next_project
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
			If next_meeting = "?" Then err_msg = err_msg & vbCr & " - Indicate if the activity is a meeting."
			If next_gh_issue <> "" and IsNumeric(next_gh_issue) = False Then err_msg = err_msg & vbCr & " - Enter the GitHub issue as the number only."

			If err_msg <> "" Then MsgBox "Need to Resolve:" & vbCr & err_msg

		Loop until err_msg = ""

		ObjExcel.Cells(current_task_row, 3).Value = end_time
		ObjExcel.Cells(current_task_row, 4).Value = "=TEXT([@[End Time]]-[@[Start Time]],"+chr(34)+"h:mm"+chr(34)+")"

		ObjExcel.Cells(next_blank_row, 1).Value = next_date
		ObjExcel.Cells(next_blank_row, 4).Value = ""
		' If next_start_time <> end_time Then ObjExcel.Cells(next_blank_row, 2).Value = next_start_time
		' ObjExcel.Cells(next_blank_row, 3).Value =
		ObjExcel.Cells(next_blank_row, 5).Value = next_category
		ObjExcel.Cells(next_blank_row, 6).Value = next_meeting
		ObjExcel.Cells(next_blank_row, 7).Value = next_detail
		ObjExcel.Cells(next_blank_row, 8).Value = "=HYPERLINK(" & chr(34) & "https://github.com/Hennepin-County/MAXIS-scripts/issues/" & next_gh_issue & chr(34) & ", " & chr(34) & next_gh_issue & chr(34) & ")"
		ObjExcel.Cells(next_blank_row, 9).Value = next_project
		ObjExcel.Cells(next_blank_row, 10).Value = "Y"
		end_msg = "As of " & next_date & " at " & end_time & " you are now working on:" & vbCr & "  - Category: " & next_category & vbCr & "  - Detail: " & next_detail

	End If

	' MsgBox "Pause"
End If



objWorkbook.Save
objWorkbook.SaveAs (t_drive_excel_file_path)
' MsgBox end_msg
ObjExcel.Quit
call script_end_procedure(end_msg)
