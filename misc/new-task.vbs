'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - New Task.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

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

EMConnect ""

' BeginDialog Dialog1, 0, 0, 296, 165, "New Task Assignment"
'   Text 10, 10, 135, 10, "Thank you for requesting a new task!"
'   Text 30, 25, 65, 10, "Case: XXXXXXX"
'   Text 50, 35, 125, 10, "(Case has been entered into MAXIS)"
'   GroupBox 5, 50, 285, 90, "Case Overview"
'   Text 15, 65, 265, 10, "Active Programs: SNAP"
'   Text 15, 80, 265, 10, "Pending Programs: Cash"
'   Text 15, 95, 65, 10, "HH Members; 2"
'   Text 15, 110, 65, 10, "DAILs found: 5"
'   Text 15, 125, 100, 10, "REVW Month: SNAP - 08/21"
'   Text 215, 10, 70, 10, "Assignment Details:"
'   Text 220, 25, 75, 10, "- Assigned on m/d/yy"
'   Text 220, 35, 70, 10, "- Assigned at hh:mm"
'   ButtonGroup ButtonPressed
'     ' CancelButton 185, 145, 50, 15
'     OkButton 240, 145, 50, 15
' EndDialog
'
' Do
' 	dialog Dialog1
' 	' cancel_without_confirmation
' Loop until ButtonPressed = -1


'--------------------------------------------------------------------------------------------'
tasks_on_hold = True
currently_on_task = false

If currently_on_task = FALSE Then
	If tasks_on_hold = TRUE Then

		task_assigned = FALSE
		Do
			BeginDialog Dialog1, 0, 0, 481, 120, "Tasks On Hold"
			  Text 25, 45, 385, 10, "Case: 654321      Started: " & date & " at 8:42 AM     Hold For: UC Verification"
			  ButtonGroup ButtonPressed
			    PushButton 415, 45, 50, 10, "Resume", resume_btn_1
			  Text 25, 60, 385, 10, "Case: 321654      Started: " & date & " at 10:15 AM     Hold For: DISQ Evaluation"
			  ButtonGroup ButtonPressed
			    PushButton 415, 60, 50, 10, "Resume", resume_btn_2
			  Text 25, 75, 385, 10, "Case: 654987      Started: " & date & " at 1:03 PM     Hold For: Knowledge Now"
			  ButtonGroup ButtonPressed
			    PushButton 415, 75, 50, 10, "Resume", resume_btn_3
			    PushButton 10, 100, 95, 15, "Start New Task", start_new_task_btn
			    ' PushButton 110, 100, 100, 15, "See all of Today's Tasks", see_task_list_btn
			    CancelButton 425, 100, 50, 15
			  Text 15, 10, 215, 10, "You currently have tasks on hold that may need to be resumed. "
			  GroupBox 10, 25, 465, 70, "On Hold Tasks"
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = start_new_task_btn Then
				BeginDialog Dialog1, 0, 0, 316, 125, "New Task Assignment"
				  ButtonGroup ButtonPressed
					OkButton 260, 105, 50, 15
					PushButton 105, 70, 35, 10, "How?", Button3
				  Text 10, 10, 120, 10, "You have been assigned to work on:"
				  Text 40, 25, 30, 10, "1234567"
				  Text 85, 40, 60, 10, "Assigned on date"
				  ' Text 85, 50, 70, 10, "Resumed from Hold"
				  GroupBox 5, 55, 150, 45, "Task Type"
				  Text 15, 70, 85, 10, "Holistic Case Processing"
				  Text 15, 85, 60, 10, "Verifications Due"
				  GroupBox 160, 5, 150, 95, "Case Overview"
				  Text 170, 20, 100, 10, "Basket: x127EN5 (Adults)"
				  Text 170, 30, 60, 10, "Active: SNAP"
				  Text 170, 40, 50, 10, "Pending: Cash"
				  Text 170, 50, 50, 10, "HH MEMBs: 2"
				  Text 170, 60, 55, 10, "DAILs Found: 5"
				  Text 170, 70, 95, 10, "REVW Month SNAP - 08/21"
				  Text 170, 80, 55, 10, "Docs in ECF: 3"
				EndDialog

				dialog Dialog1
				task_assigned = TRUE
			End If
			If ButtonPressed = see_task_list_btn Then Call list_of_all_tasks
			If ButtonPressed = resume_btn_1 Then
				MsgBox "You have resumed your task on Case # 654321"
				task_assigned = TRUE
			End If
			If ButtonPressed = resume_btn_2 Then
				MsgBox "You have resumed your task on Case # 321654"
				task_assigned = TRUE
			End If
			If ButtonPressed = resume_btn_3 Then
				MsgBox "You have resumed your task on Case # 654987"
				task_assigned = TRUE
			End If


		Loop until task_assigned = TRUE

	Else
		BeginDialog Dialog1, 0, 0, 316, 125, "New Task Assignment"
		  ButtonGroup ButtonPressed
			OkButton 260, 105, 50, 15
			PushButton 105, 70, 35, 10, "How?", Button3
		  Text 10, 10, 120, 10, "You have been assigned to work on:"
		  Text 40, 25, 30, 10, "1234567"
		  Text 85, 40, 60, 10, "Assigned on date"
		  ' Text 85, 50, 70, 10, "Resumed from Hold"
		  GroupBox 5, 55, 150, 45, "Task Type"
		  Text 15, 70, 85, 10, "Holistic Case Processing"
		  Text 15, 85, 60, 10, "Verifications Due"
		  GroupBox 160, 5, 150, 95, "Case Overview"
		  Text 170, 20, 100, 10, "Basket: x127EN5 (Adults)"
		  Text 170, 30, 60, 10, "Active: SNAP"
		  Text 170, 40, 50, 10, "Pending: Cash"
		  Text 170, 50, 50, 10, "HH MEMBs: 2"
		  Text 170, 60, 55, 10, "DAILs Found: 5"
		  Text 170, 70, 95, 10, "REVW Month SNAP - 08/21"
		  Text 170, 80, 55, 10, "Docs in ECF: 3"
		EndDialog

		dialog Dialog1
	End If
Else
	MsgBox "You are currently on a task (Case # XXXXXX) and you cannot start a new task." & vbCr & vbCr & "You must log your current task first."
End If

call script_end_procedure("")
