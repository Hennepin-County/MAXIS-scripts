'Required for statistical purposes==========================================================================================
name_of_script = "TASKS - Complete Task.vbs"
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
' BeginDialog Dialog1, 0, 0, 336, 210, "Task Completion Information"
'   DropListBox 170, 20, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List2
'   DropListBox 170, 35, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List3
'   DropListBox 15, 65, 180, 45, "No follow up needed"+chr(9)+"Yes - policy and process questions"+chr(9)+"Yes - Unique/specific task assignment"+chr(9)+"Yes - Assingment/work question", List1
'   CheckBox 20, 100, 70, 10, "MFIP Sanctions", Check1
'   CheckBox 95, 100, 70, 10, "Immigration", Check3
'   CheckBox 175, 100, 70, 10, "More...", Check5
'   CheckBox 20, 115, 70, 10, "Facility", Check2
'   CheckBox 95, 115, 70, 10, "Overpayment", Check4
'   TextBox 175, 115, 100, 15, more_detail_var
'   DropListBox 130, 130, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List4
'   DropListBox 130, 145, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List5
'   DropListBox 130, 160, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List6
'   DropListBox 140, 175, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", List7
'   CheckBox 15, 195, 90, 10, "Send updates to METS", Check6
'   ButtonGroup ButtonPressed
'     ' CancelButton 225, 190, 50, 15
'     OkButton 280, 190, 50, 15
'   Text 5, 10, 135, 10, "Task Completion for Case: XXXXXXX"
'   Text 10, 25, 155, 10, "Was there work to be completed on this case?"
'   Text 10, 40, 150, 10, "Were you able to complete all the work?"
'   Text 10, 55, 105, 10, "Does this case need follow up?"
'   Text 15, 85, 290, 10, "If this task should be handled by a  specialty group, select the appropriate actions here:"
'   Text 15, 135, 110, 10, "Did you complete an interview?"
'   Text 15, 150, 110, 10, "Did you 'APP' in ELIG?"
'   Text 15, 165, 110, 10, "Did you CASE:NOTE?"
'   Text 15, 180, 125, 10, "Did you send ECF Docs to the client?"
'   Text 245, 10, 70, 10, "Assignment Details:"
'   Text 250, 25, 75, 10, "- Completed on m/d/yy"
'   Text 250, 35, 80, 10, "- Completed at hh:mm"
'   ' Text 250, 45, 80, 10, "- Time Spent - h:mm"
' EndDialog
' Do
' 	dialog Dialog1
' 	' cancel_without_confirmation
' Loop until ButtonPressed = -1

'------------------------------------------------------------------------------------------------'
currently_on_task = true

complete_task = TRUE
hold_task = FALSE
If currently_on_task = TRUE Then
	Do
		If complete_task = TRUE Then
			BeginDialog Dialog1, 0, 0, 336, 220, "Log Task Information"
			  DropListBox 235, 35, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", work_to_be_completed_ans
			  DropListBox 235, 50, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", all_work_completed_ans
			  DropListBox 120, 90, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_intv_ans
			  DropListBox 120, 105, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_APP_ans
			  DropListBox 120, 120, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_accept_ecf_ans
			  DropListBox 120, 135, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", did_you_send_ecf_ans
			  DropListBox 275, 150, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", start_cic_ans
			  DropListBox 275, 165, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", more_processing_ans
			  GroupBox 10, 15, 315, 180, ""
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
				PushButton 225, 175, 90, 15, "Save Task Log", save_task_log_btn
				' PushButton 20, 10, 70, 15, "Complete Task", complete_task_btn
				Text 30, 13, 65, 10, "Complete Task"
				PushButton 95, 10, 120, 15, "Request Support - Task On Hold", hold_task_btn
				' PushButton 10, 200, 100, 15, "See all of Today's Tasks", see_task_list_btn
				CancelButton 280, 200, 50, 15
			EndDialog

		ElseIf hold_task = True Then
			BeginDialog Dialog1, 0, 0, 336, 220, "Log Task Information"
			  GroupBox 10, 15, 315, 180, ""
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
				PushButton 225, 175, 90, 15, "Save Task Log", save_task_log_btn
				PushButton 20, 10, 70, 15, "Complete Task", complete_task_btn
				Text 100, 13, 115, 10, "Request Support - Task On Hold"
				' PushButton 95, 10, 120, 15, "Request Support - Task On Hold", hold_task_btn
				' PushButton 10, 200, 100, 15, "See all of Today's Tasks", see_task_list_btn
				CancelButton 280, 200, 50, 15
			EndDialog

		End If

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed = see_task_list_btn Then Call list_of_all_tasks

		If ButtonPressed = complete_task_btn Then
			complete_task = TRUE
			hold_task = FALSE
		End If
		If ButtonPressed = hold_task_btn Then
			complete_task = FALSE
			hold_task = TRUE
		End If

	Loop until ButtonPressed = save_task_log_btn
Else
	MsgBox "You are not currently on a task, so there is nothing to log." & vbCr & vbCR & "You can either:" & vbCr & " - Start a new Task" & vbCr & " - Resume an ON HOLD task" & vbCr & " - Select a completed task from the list to update the log."
End If

call script_end_procedure("")
