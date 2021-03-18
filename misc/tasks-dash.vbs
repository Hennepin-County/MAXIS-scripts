'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - DASH.vbs"
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

'
'


tasks_on_hold = TRUE
currently_on_task = FALSE
phone_shift = FALSE
running_as_HSS = FALSE



ssr_worker = FALSE
qi_worker = FALSE

start_new_task_btn	= 100
hsr_manual_btn		= 101
save_task_log_btn	= 102
complete_task_btn	= 103
hold_task_btn		= 104
see_task_list_btn	= 105
resume_btn_1		= 106
resume_btn_2		= 107
resume_btn_3		= 108
update_log_btn		= 109
log_task_btn		= 110

function list_of_all_tasks()
	date_to_review = date & ""
	date_change = date_to_review
	Do
		BeginDialog Dialog1, 0, 0, 486, 230, "All Tasks Assigned"
		  ButtonGroup ButtonPressed
		    CancelButton 430, 210, 50, 15
		  Text 10, 10, 110, 10, "All tasks for WORKER NAME"
		  If currently_on_task = TRUE Then Text 125, 10, 200, 10, "***** CURRENTLY ON TASK *****  Case # XXXXXX"

		  GroupBox 10, 25, 470, 65, "Tasks On Hold on " & date_to_review
		  Text 20, 40, 35, 10, "Case #"
		  Text 80, 40, 40, 10, "Hold Start"
		  Text 170, 40, 50, 10, "Hold Reason"
		  Text 20, 55, 40, 10, "654321"
		  Text 80, 55, 50, 10, "8:42 AM"
		  Text 170, 55, 205, 10, "UC Verification"
		  Text 20, 65, 40, 10, "321654"
		  Text 80, 65, 50, 10, "10:15 AM"
		  Text 170, 65, 205, 10, "DISQ Evaluation"
		  Text 20, 75, 40, 10, "654987"
		  Text 80, 75, 50, 10, "1:03 PM"
		  Text 170, 75, 205, 10, "Knowledge Now"

		  GroupBox 10, 95, 470, 110, "Tasks Commpleted on " & date_to_review
		  Text 20, 110, 35, 10, "Case #"
		  Text 80, 110, 40, 10, "Time Start"
		  Text 190, 110, 20, 10, "APP"
		  Text 135, 110, 35, 10, "Interview "
		  Text 225, 110, 50, 10, "ECF Doc Sent"
		  Text 295, 110, 50, 10, "Parallel Task"
		  Text 80, 125, 35, 10, "7:05 AM"
		  Text 135, 125, 25, 10, "No"
		  Text 190, 125, 25, 10, "Yes"
		  Text 225, 125, 25, 10, "No"
		  Text 295, 125, 115, 10, "Yes"
		  CheckBox 20, 125, 50, 10, "XXXXXX", Check1
		  Text 80, 135, 35, 10, "8:20 AM"
		  Text 135, 135, 25, 10, "No"
		  Text 190, 135, 25, 10, "Yes"
		  Text 225, 135, 25, 10, "No"
		  Text 295, 135, 115, 10, "No"
		  CheckBox 20, 135, 50, 10, "XXXXXX", Check2
		  Text 80, 145, 35, 10, "8:45 AM"
		  Text 135, 145, 25, 10, "No"
		  Text 190, 145, 25, 10, "No"
		  Text 225, 145, 25, 10, "Yes"
		  Text 295, 145, 115, 10, "No"
		  CheckBox 20, 145, 50, 10, "XXXXXX", Check3
		  Text 80, 155, 35, 10, "8:55 AM"
		  Text 135, 155, 25, 10, "Yes"
		  Text 190, 155, 25, 10, "No"
		  Text 225, 155, 25, 10, "Yes"
		  Text 295, 155, 115, 10, "No"
		  CheckBox 20, 155, 50, 10, "XXXXXX", Check4
		  Text 80, 165, 35, 10, "9:34 AM"
		  Text 135, 165, 25, 10, "Yes"
		  Text 190, 165, 25, 10, "Yes"
		  Text 225, 165, 25, 10, "Yes"
		  Text 295, 165, 115, 10, "No"
		  CheckBox 20, 165, 50, 10, "XXXXXX", Check5
		  Text 80, 175, 35, 10, "10:27 AM"
		  Text 135, 175, 25, 10, "No"
		  Text 190, 175, 25, 10, "No"
		  Text 225, 175, 25, 10, "Yes"
		  Text 295, 175, 115, 10, "Yes"
		  CheckBox 20, 175, 50, 10, "XXXXXX", Check6
		  Text 80, 185, 35, 10, "11:03 AM"
		  Text 135, 185, 25, 10, "Yes"
		  Text 190, 185, 25, 10, "Yes"
		  Text 225, 185, 25, 10, "No"
		  Text 295, 185, 115, 10, "No"
		  CheckBox 20, 185, 50, 10, "XXXXXX", Check7
		  Text 320, 10, 55, 10, "Date to Review:"
		  EditBox 375, 5, 50, 15, date_change
		  ButtonGroup ButtonPressed
		    PushButton 430, 10, 50, 10, "CHANGE", change_date_btn
			PushButton 390, 110, 85, 10, "Update 'checked' logs", update_log_btn
		    If currently_on_task = FALSE Then
				PushButton 390, 55, 50, 10, "Resume", resume_btn_1
				PushButton 390, 65, 50, 10, "Resume", resume_btn_2
				PushButton 390, 75, 50, 10, "Resume", resume_btn_3
				PushButton 10, 210, 95, 15, "Start New Task", start_new_task_btn
			End If
			If currently_on_task = TRUE Then PushButton 10, 210, 95, 15, "Log Task", log_task_btn
		EndDialog

		dialog Dialog1

		cancel_without_confirmation

		If ButtonPressed = change_date_btn Then
			date_to_review = date_change
		End If

	Loop until ButtonPressed <> change_date_btn



	If ButtonPressed = start_new_task_btn Then
		call assign_a_task
		task_assigned = TRUE
	End If

	If ButtonPressed = update_log_btn Then
		MsgBox "Log detail for the selected case would open here."
		script_end_procedure("")
	End If

end function

function assign_a_task()
	Do
		BeginDialog Dialog1, 0, 0, 316, 125, "New Task Assignment"
		  ButtonGroup ButtonPressed
		    OkButton 260, 105, 50, 15
		    PushButton 105, 70, 35, 10, "How?", hsr_manual_btn
			PushButton 10, 105, 100, 15, "See all of Today's Tasks", see_task_list_btn
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
		cancel_without_confirmation

		currently_on_task = True

		If ButtonPressed = hsr_manual_btn Then MsgBox "This will link to an HSR manual page that details 'Holistic Processing'"
		If ButtonPressed = see_task_list_btn Then Call list_of_all_tasks
	Loop Until ButtonPressed <> hsr_manual_btn
end function

If running_as_HSS = FALSE Then
	If currently_on_task = TRUE Then
		complete_task = TRUE
		hold_task = FALSE

		Do
			If complete_task = TRUE Then
				BeginDialog Dialog1, 0, 0, 336, 230, "Log Task Information"
				  DropListBox 235, 35, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", work_to_be_completed_ans
				  DropListBox 235, 50, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", all_work_completed_ans
				  DropListBox 120, 90, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_intv_ans
				  DropListBox 120, 105, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_APP_ans
				  DropListBox 120, 120, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_accept_ecf_ans
				  DropListBox 120, 135, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_send_ecf_ans
				  DropListBox 275, 150, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", start_cic_ans
				  DropListBox 275, 165, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", more_processing_ans
				  GroupBox 10, 15, 315, 190, ""
				  Text 25, 35, 85, 10, "Case: XXXXXXX"
				  Text 115, 40, 115, 10, "Was there work to be completed?"
				  Text 35, 55, 195, 10, "Were you able to complete EVERYTHING the case needs?"
				  Text 25, 80, 50, 10, "Did you ..."
				  Text 35, 95, 80, 10, "Complete an Interview?"
				  Text 55, 110, 60, 10, "APP a Pprogram?"
				  Text 35, 125, 85, 10, "Accept ECF Documents?"
				  Text 20, 140, 100, 10, "Send a Document from ECF?"
				  Text 145, 155, 125, 10, "Do you need to start the CIC process?"
				  Text 65, 170, 210, 10, "Does this case need additional processing from another group?"
				  ButtonGroup ButtonPressed
				    PushButton 230, 185, 90, 15, "Save Task Log", save_task_log_btn
				    ' PushButton 20, 10, 70, 15, "Complete Task", complete_task_btn
				    Text 30, 10, 65, 10, "Complete Task"
				    PushButton 95, 10, 120, 15, "Request Support - Task On Hold", hold_task_btn
				    PushButton 10, 210, 100, 15, "See all of Today's Tasks", see_task_list_btn
				    CancelButton 280, 210, 50, 15
				EndDialog

			ElseIf hold_task = True Then
				BeginDialog Dialog1, 0, 0, 336, 230, "Log Task Information"
				  GroupBox 10, 15, 315, 190, ""
				  Text 25, 35, 85, 10, "Case: XXXXXXX"
				  Text 25, 50, 120, 10, "What kind of support do you need?"
				  DropListBox 145, 45, 175, 45, "Select One..."+chr(9)+"Policy and Procedure Questions (Knowledge Now)"+chr(9)+"Request to APPL"+chr(9)+"Verifications"+chr(9)+"Imigration Review"+chr(9)+"Script Error"+chr(9)+"HSS Review/Approval"+chr(9)+"Fresh Eyes Request"+chr(9)+"Lost ApplyMN"+chr(9)+"DISQ Evaluation for Removal", hold_type

				  If hold_type = "Policy and Procedure Questions (Knowledge Now)" Then
					  Text 25, 65, 105, 10, "Programs you need support on:"
					  CheckBox 25, 75, 30, 10, "SNAP", Check1
					  CheckBox 60, 75, 30, 10, "MFIP", Check2
					  CheckBox 95, 75, 30, 10, "DWP", Check3
					  CheckBox 125, 75, 20, 10, "GA", Check4
					  CheckBox 150, 75, 25, 10, "MSA", Check5
					  CheckBox 180, 75, 40, 10, "GRH/HS", Check6
					  CheckBox 220, 75, 25, 10, "HC", Check7
					  CheckBox 245, 75, 35, 10, "EA/EGA", Check8
					  Text 25, 95, 40, 10, "Case type:"
					  DropListBox 65, 90, 60, 45, "Select"+chr(9)+"Adults"+chr(9)+"Families", List10
					  Text 25, 110, 125, 10, "Explain what you need support with:"
					  EditBox 25, 120, 295, 15, Edit1
					  Text 25, 140, 125, 10, "Additional notes on this case:"
					  EditBox 25, 150, 295, 15, Edit5
				  End If
				  If hold_type = "Request to APPL" Then
					  Text 25, 70, 130, 10, "What case situation needs pending?"
					  DropListBox 25, 80, 295, 45, "Select One..."+chr(9)+"CAF in ECF - programm not pending in MAXIS"+chr(9)+"MA Transition"+chr(9)+"Auto-Newborn Retro MA", List11
					  Text 25, 105, 40, 10, "APPL Date:"
					  EditBox 65, 100, 50, 15, Edit7
					  Text 25, 125, 160, 10, "Is the client on the phone or awaiting call-back?"
					  DropListBox 185, 120, 40, 45, "?"+chr(9)+"No"+chr(9)+"Yes", List12
				  End If
				  If hold_type = "Verifications" Then
					  Text 25, 70, 105, 10, "What Verification is Needed?"
					  CheckBox 30, 90, 50, 10, "UC Income", Check9
					  CheckBox 30, 110, 65, 10, "VA Information", Check10
					  CheckBox 30, 130, 65, 10, "SSA Information", Check11
					  Text 100, 90, 45, 10, "SSNs to run:"
					  EditBox 145, 85, 175, 15, Edit8
					  Text 100, 110, 45, 10, "SSNs to run:"
					  EditBox 145, 105, 175, 15, Edit9
					  Text 100, 130, 45, 10, "SSNs to run:"
					  EditBox 145, 125, 175, 15, Edit10
				  End If
				  If hold_type = "Immigration Review" Then

				  End If
				  If hold_type = "Script Error" Then

				  End If
				  If hold_type = "HSS Review/Approval" Then

				  End If
				  If hold_type = "Fresh Eyes Request" Then

				  End If
				  If hold_type = "Lost ApplyMN" Then

				  End If
				  If hold_type = "DISQ Evaluation for Removal" Then
				  End If

				  ButtonGroup ButtonPressed
					PushButton 270, 60, 50, 10, "Add Details", add_details_btn
					PushButton 230, 185, 90, 15, "Save Task Log", save_task_log_btn
					PushButton 20, 10, 70, 15, "Complete Task", complete_task_btn
					Text 100, 13, 115, 10, "Request Support - Task On Hold"
					' PushButton 95, 10, 120, 15, "Request Support - Task On Hold", hold_task_btn
					PushButton 10, 210, 100, 15, "See all of Today's Tasks", see_task_list_btn
	  			    CancelButton 280, 210, 50, 15
				EndDialog

			End If

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = see_task_list_btn Then
				Call list_of_all_tasks
				If ButtonPressed <> log_task_btn Then script_end_procedure("")
			End If

			If ButtonPressed = save_task_log_btn AND complete_task = TRUE Then
				If all_work_completed_ans = "No" OR more_processing_ans = "Yes" Then
					Do
						BeginDialog Dialog1, 0, 0, 171, 140, "Passing the Baton"
						  CheckBox 15, 65, 65, 10, "Overpayments", overpayment_check
						  CheckBox 15, 80, 65, 10, "EA/EGA", emer_check
						  CheckBox 15, 95, 65, 10, "GRH", grh_check
						  CheckBox 15, 110, 85, 10, "Issue MONY/CHCK", mony_chck_check
						  ButtonGroup ButtonPressed
						    PushButton 70, 125, 95, 10, "Add Details", add_dtls_btn
						  Text 60, 10, 75, 10, "Case # XXXXXX"
						  Text 10, 25, 160, 20, "You have indicated that this case has additional work to be completed by another group."
						  Text 10, 50, 155, 10, "Check each work type that this case requires:"
						EndDialog

						dialog Dialog1

						If overpayment_check = Checked Then

							BeginDialog Dialog1, 0, 0, 336, 155, "Passing the Baton"
							  EditBox 15, 50, 310, 15, Edit1
							  EditBox 15, 80, 310, 15, Edit2
							  EditBox 15, 110, 310, 15, Edit4
							  ButtonGroup ButtonPressed
							    PushButton 215, 140, 115, 10, "Submit Reassignment Task", Button3
							  Text 140, 10, 75, 10, "Case # XXXXXX"
							  GroupBox 10, 25, 320, 110, "Overpayments Processing"
							  Text 15, 40, 145, 10, "List Programs with Potential Overpayments:"
							  Text 15, 70, 145, 10, "All Potential Months of Overpayment:"
							  Text 15, 100, 145, 10, "Overpayment Reasons and Notes:"
							EndDialog

							dialog Dialog1

						End If

						reassign_msg = "You have successfully reassigned this case for processing on:" & vbCr

						If overpayment_check = Checked Then reassign_msg = reassign_msg & vbCr & " - Overpayments"
						If emer_check = Checked Then reassign_msg = reassign_msg & vbCr & " - EA/EGA"
						If grh_check = Checked Then reassign_msg = reassign_msg & vbCr & " - GRHH"
						If mony_chck_check = Checked Then reassign_msg = reassign_msg & vbCr & " - MONY/CHCK"

						If overpayment_check = unchecked AND emer_check = unchecked AND grh_check = unchecked AND mony_chck_check = unchecked Then reassign_msg = "No reassignment tasks were selected."

						confirm_reassignment = MsgBox(reassign_msg & vbCr & vbCr & "Is this the correct reassignment action?", vbYesNo, "Confirm Reassgnment")
					Loop until confirm_reassignment = vbYes
					ButtonPressed = save_task_log_btn
				End If
				MsgBox "Your Task has been Commpleted and Logged"
			End If

			If ButtonPressed = save_task_log_btn AND hold_task = True Then
				MsgBox "Your Task has been put ON HOLD."
			End If
			If ButtonPressed = complete_task_btn Then
				complete_task = TRUE
				hold_task = FALSE
			End If
			If ButtonPressed = hold_task_btn Then
				complete_task = FALSE
				hold_task = TRUE
			End If

		Loop until ButtonPressed = save_task_log_btn

	End If

	If tasks_on_hold = TRUE and currently_on_task = FALSE Then

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
			    PushButton 110, 100, 100, 15, "See all of Today's Tasks", see_task_list_btn
			    CancelButton 425, 100, 50, 15
			  Text 15, 10, 215, 10, "You currently have tasks on hold that may need to be resumed. "
			  GroupBox 10, 25, 465, 70, "On Hold Tasks"
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = start_new_task_btn Then
				call assign_a_task
				task_assigned = TRUE
			End If
			If ButtonPressed = see_task_list_btn Then
				Call list_of_all_tasks
				' If ButtonPressed <> start_new_task_btn Then script_end_procedure("")
			End If
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

	End If

	If currently_on_task = FALSE AND tasks_on_hold = FALSE Then call assign_a_task

End If

If running_as_HSS = TRUE Then
	BeginDialog Dialog1, 0, 0, 291, 205, "HSS DASH Board"
	  ButtonGroup ButtonPressed
	    PushButton 20, 50, 125, 15, "Tasks for a Individual Worker",Button1
	    PushButton 150, 50, 125, 15, "Tasks by Type", Button9
	    PushButton 20, 70, 125, 15, "Team Task Counts", Button5
	    PushButton 150, 70, 125, 15, "Workers not on a Task", Button11
	    PushButton 20, 90, 125, 15, "My Input Task Lists", Button7
	    PushButton 150, 90, 125, 15, "Pending Cases", Button14
	    PushButton 20, 135, 125, 15, "Add List of Tasks", Button17
	    PushButton 150, 135, 125, 15, "Reserve a Task for a Worker", Button21
	    PushButton 20, 155, 125, 15, "Unassign a Task", Button19
	    PushButton 150, 155, 125, 15, "Create a Task", Button22
	    PushButton 10, 185, 130, 15, "Update HSS Assignment for the Day", Button25
	    CancelButton 235, 185, 50, 15
	  Text 10, 10, 270, 20, "Welcome to the DASH Board for the HSS User Role. This menu provides access to the functionnality availale to HSSs."
	  GroupBox 10, 35, 275, 80, "Task Review"
	  GroupBox 10, 120, 275, 60, "Task Updates"
	EndDialog

	dialog Dialog1

End If

script_end_procedure("")
