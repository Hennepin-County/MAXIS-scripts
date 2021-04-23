'Required for statistical purposes==========================================================================================
name_of_script = "I am a test.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
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


' function script_end_procedure_with_error_report(closing_message)
' '--- This function is how all user stats are collected when a script ends.
' '~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
' '===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
' 	stop_time = timer
'     send_error_message = ""
' 	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
'         send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbDefaultButton2 + vbYesNo, "Script Run Completed")
'     End If
'     script_run_time = stop_time - start_time
' 	If is_county_collecting_stats  = True then
' 		'Getting user name
' 		Set objNet = CreateObject("WScript.NetWork")
' 		user_ID = objNet.UserName
'
' 		'Setting constants
' 		Const adOpenStatic = 3
' 		Const adLockOptimistic = 3
'
'         'Determining if the script was successful
'         If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
'             SCRIPT_success = -1
'         else
'             SCRIPT_success = 0
'         end if
'
' 		'Determines if the value of the MAXIS case number - BULK and UTILITIES scripts will not have case number informaiton input into the database
' 		IF left(name_of_script, 4) = "BULK" or left(name_of_script, 4) = "UTIL" then
' 			MAXIS_CASE_NUMBER = ""
' 		End if
'
' 		'Creating objects for Access
' 		Set objConnection = CreateObject("ADODB.Connection")
' 		Set objRecordSet = CreateObject("ADODB.Recordset")
'
' 		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
' 		closing_message = replace(closing_message, "'", "")
'
' 		'Opening DB
' 		IF using_SQL_database = TRUE then
'     		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
' 		ELSE
' 			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
' 		END IF
'
'         'Adds some data for users of the old database, but adds lots more data for users of the new.
'         If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
'     		'Opening usage_log and adding a record
'     		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
'     		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
' 		'collecting case numbers counties
' 		Elseif collect_MAXIS_case_number = true then
' 			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
' 			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
' 		 'for users of the new db
' 		Else
'             objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
'             "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
'         End if
'
' 		'Closing the connection
' 		objConnection.Close
' 	End if
'
'     If send_error_message = vbYes Then
'         'dialog here to gather more detail
'         Dialog1 = ""
'         BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
'           Text 60, 35, 55, 10, MAXIS_case_number
'           ComboBox 220, 30, 175, 45, ""+chr(9)+"BUG - somethng happened that was wrong"+chr(9)+"ENHANCEMENT - somthing could be done better"+chr(9)+"TYPO - gramatical/spelling type errors", error_type
'           EditBox 65, 50, 330, 15, error_detail
'           CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
'           CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
'           CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
'           CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
'           CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
'           CheckBox 95, 115, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
'           CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
'           CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
'           EditBox 60, 140, 165, 15, worker_signature
'           ButtonGroup ButtonPressed
'             OkButton 290, 140, 50, 15
'             CancelButton 345, 140, 50, 15
'           Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
'           Text 5, 35, 50, 10, "Case Number:"
'           Text 125, 35, 95, 10, "What type of error occured?"
'           Text 5, 55, 60, 10, "Explain in detail:"
'           GroupBox 10, 75, 380, 60, "Common areas of issue"
'           Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
'           Text 10, 145, 50, 10, "Worker Name:"
'           Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
'         EndDialog
'
'         Dialog Dialog1
'
'         'sent email here
'         If ButtonPressed = -1 Then
'             bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
'             subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"
'
'             full_text = "Error occured on " & date & " at " & time
'             full_text = full_text & vbCr & "Error type - " & error_type
'             full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
'             full_text = full_text & vbCr & "Information: " & error_detail
'             If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"
'
'             If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
'             If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
'             If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
'             If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
'             If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
'             If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
'             If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
'             If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"
'
'             full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature
'
'             If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown
'
'             Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)
'
'             MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
'         Else
'             MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
'         End If
'     End If
' 	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
' end function


' Call MAXIS_case_number_finder(MAXIS_case_number)
'
' Dialog1 = ""
' BeginDialog Dialog1, 0, 0, 126, 55, "Dialog"
'   ButtonGroup ButtonPressed
'     OkButton 70, 30, 50, 15
'   Text 10, 15, 50, 10, "Case Number"
'   EditBox 60, 10, 60, 15, MAXIS_case_number
' EndDialog
'
' Do
'     err_msg = ""
'
'     dialog Dialog1
'     Call validate_MAXIS_case_number(err_msg, "-")
'     If err_msg <> "" Then MsgBox("Please review the foloowing in order for the script to continue:" & vbNewLine & err_msg)
'
' Loop until err_msg = ""
'
' Call start_a_blank_CASE_NOTE
' notes_variable = "03/19 for 01 is BANKED MONTH - Banked Month: 3.; 04/19 for 01 is BANKED MONTH - Banked Month: 4.;"
' bullet_variable = "This is where the bullet would be all the things."
' time_variable = "Now is the time and this is the place."
' order_variable = "Everything in it's place."
'
' Call write_variable_in_CASE_NOTE("*** SNAP approved starting in 03/19 ***")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 03/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 04/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_bullet_and_variable_in_CASE_NOTE("Notes", notes_variable)
' Call write_variable_in_CASE_NOTE("This is a thing")
' Call write_variable_in_CASE_NOTE("   this is another thing")
' Call write_variable_in_CASE_NOTE("How now brown cow")
' Call write_variable_in_CASE_NOTE("the thing and thing and stuff")
' Call write_variable_in_CASE_NOTE("all the writing")
' Call write_variable_in_CASE_NOTE("blah blah blah")
' Call write_bullet_and_variable_in_CASE_NOTE("BULLET", bullet_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Time", time_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Order", order_variable)
'
' Call write_variable_in_CASE_NOTE("H.Lamb/QI")

' script_list_URL = "C:\MAXIS-scripts\Test scripts\Casey\User Group\COMPLETE LIST OF TESTERS.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

' Call confirm_tester_information


function MFIP_cert_length_details(verbal_attestation, attestation_verif_array)
' This script requires
'~~~~~ verbal_attestation: BOOLEAN - that idetifies if ANY verbal attestation was used

	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
	ReDim attestation_verif_array(0)
	If mfip_case = TRUE Then
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		EMReadScreen ER_Month, 2, 9, 37
		EMReadScreen ER_Year, 2, 9, 43
		If snap_case = TRUE Then
			EmWriteScreen "X", 5, 58
			transmit
			EMReadScreen SNAP_ER_Month, 2, 9, 64
			EMReadScreen SNAP_ER_Year, 2, 9, 70
			transmit
		End If

		Do
			err_msg = ""
			If verif_by_attestation_yn = "" Then
				BeginDialog Dialog1, 0, 0, 256, 30, "Verification by Attestation Details"
			Else
				dlg_len = 155
				If attestation_verif_array(0) <> "" Then dlg_len = dlg_len + ((UBound(attestation_verif_array)+1) * 15)
				If verbal_attestation = FALSE Then dlg_len = 100

				BeginDialog Dialog1, 0, 0, 256, dlg_len, "Verification by Attestation Details"
			End If

			  ButtonGroup ButtonPressed
			    Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
			    DropListBox 155, 10, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", verif_by_attestation_yn
			    If verif_by_attestation_yn = "" Then PushButton 210, 10, 40, 15, "Enter", enter_btn
				If verbal_attestation = TRUE Then
					Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
				    ' Text 20, 60, 235, 10, "VERIF DETIL HERE"

					y_pos = 60
					If attestation_verif_array(0) <> "" Then
						For each item in attestation_verif_array
							Text 20, y_pos, 235, 10, "- " & item
							y_pos = y_pos + 15
						Next
					End If

					Text 15, y_pos, 90, 10, "Enter a single verification:"
					y_pos = y_pos + 10
					EditBox 15, y_pos, 230, 15, verif_entry
					y_pos = y_pos + 20
				    PushButton 145, y_pos, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
					y_pos = y_pos + 25
				    Text 10, y_pos, 125, 10, "Current or Upcoming Renewal Month:"
				    EditBox 140, y_pos - 5, 15, 15, ER_Month
				    EditBox 160, y_pos - 5, 15, 15, ER_Year
				    Text 180, y_pos, 30, 10, "(MM YY)"
					y_pos = y_pos + 20
					PushButton 120, y_pos, 130, 15, "Finish Saving Attestation Information", finish_btn
				End If

				If verbal_attestation = FALSE Then
					Text 10, 35, 225, 20, "Since all verifications have been received according to 'non-waiver' policy, confirm the renewal month."
				    Text 10, 60, 125, 10, "Current or Upcoming Renewal Month:"
				    EditBox 140, 55, 15, 15, ER_Month
				    EditBox 160, 55, 15, 15, ER_Year
				    Text 180, 60, 30, 10, "(MM YY)"
				    PushButton 120, 80, 130, 15, "Finish Cert Period Assesment", finish_btn
				End If
			EndDialog

			Dialog Dialog1

			If verif_by_attestation_yn = "Yes" Then verbal_attestation = TRUE
			If verif_by_attestation_yn = "No" Then verbal_attestation = FALSE

			verif_entry = trim(verif_entry)
			If verif_entry <> "" Then
				If attestation_verif_array(0) = "" Then
					next_verif = 0
				Else
					next_verif = UBound(attestation_verif_array) + 1
					ReDim Preserve attestation_verif_array(next_verif)
				End If
				attestation_verif_array(next_verif) = verif_entry
				verif_entry = ""
			End If

			If ButtonPressed <> finish_btn Then err_msg = "LOOP"

		Loop until err_msg = ""

		how_far_away_is_the_next_REVW = ""
		next_REVW_date = ER_Month & "/1/" & ER_Year
		next_REVW_date = DateAdd("d", 0, next_REVW_date)

		how_far_away_is_the_next_REVW = DateDiff("m", date, next_REVW_date)
		MsgBox how_far_away_is_the_next_REVW

		before_er_cutoff = FALSE
		current_day = DatePart("d", date)
		If current_day < 16 Then before_er_cutoff = TRUE

		If verbal_attestation = FALSE AND how_far_away_is_the_next_REVW < 8 Then

			Call Navigate_to_MAXIS_screen("CASE", "NOTE")


			first_day_of_this_process = #4/15/2021#

            note_row = 5        'these always need to be reset when looking at Case note
            note_date = ""
            note_title = ""
			previously_set_to_6_months = FALSE			
            Do                  'this do-loop moves down the list of case notes - looking at each row in MAXIS
                EMReadScreen note_date, 8, note_row, 6      'reading the date of the row
                EMReadScreen note_title, 55, note_row, 25   'reading the header of the note
                note_title = trim(note_title)               'trim it down

                'if the note headers match any of the following then we can know if a face to face is needed or not - then we add that detail to the ARRAY
                If trim(note_title) = "MFIP Certification Period set for 6 MONTHS due Verification by Attestation" Then
					previously_set_to_6_months = TRUE
					Exit Do
				End If
				If trim(note_title) = "MFIP Certification Period set for 12 MONTHS since Verifs have been Received" Then Exit Do

                IF note_date = "        " then Exit Do      'if the case is new, we will hit blank note dates and we don't need to read any further
                note_row = note_row + 1                     'going to the next row to look at the next notws
                IF note_row = 19 THEN                       'if we have reached the end of the list of case notes then we will go to the enxt page of notes
                    PF8
                    note_row = 5
                END IF
                EMReadScreen next_note_date, 8, note_row, 6 'looking at the next note date
                IF next_note_date = "        " then Exit Do
            Loop until datevalue(next_note_date) < first_day_of_this_process 'looking ahead at the next case note kicking out the dates before app'

		End If

		Call start_a_blank_CASE_NOTE
		If verbal_attestation = TRUE Then
			Call write_variable_in_CASE_NOTE("MFIP Certification Period set for 6 MONTHS due Verification by Attestation")
			Call write_variable_in_CASE_NOTE()
			Call write_variable_in_CASE_NOTE()
			Call write_variable_in_CASE_NOTE("---")
			Call write_variable_in_CASE_NOTE(worker_signature)
		End If

		If verbal_attestation = FALSE Then
			Call write_variable_in_CASE_NOTE("MFIP Certification Period set for 12 MONTHS since Verifs have been Received")
			Call write_variable_in_CASE_NOTE()
			Call write_variable_in_CASE_NOTE()
			Call write_variable_in_CASE_NOTE("---")
			Call write_variable_in_CASE_NOTE(worker_signature)

		End If

	End If


end function


BeginDialog Dialog1, 0, 0, 256, 170, "Verbal Attestation Details"
  Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
  DropListBox 150, 10, 40, 45, "", verif_by_attestation_yn
  Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
  Text 20, 60, 235, 10, "VERIF DETIL HERE"
  Text 15, 75, 90, 10, "Enter a single verification:"
  EditBox 15, 85, 230, 15, Edit1
  ButtonGroup ButtonPressed
    PushButton 145, 105, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
  Text 10, 130, 125, 10, "Current or Upcoming Renewal Month:"
  EditBox 140, 125, 15, 15, Edit2
  EditBox 160, 125, 15, 15, Edit3
  Text 180, 130, 30, 10, "(MM YY)"
  ButtonGroup ButtonPressed
    PushButton 95, 150, 155, 15, "Finish Saving Verbal Attestation Information", finish_btn
EndDialog

BeginDialog Dialog1, 0, 0, 256, 155, "Verbal Attestation Details"
  Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
  DropListBox 150, 10, 40, 45, "", verif_by_attestation_yn
  Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
  Text 15, 60, 90, 10, "Enter a single verification:"
  EditBox 15, 70, 230, 15, Edit1
  ButtonGroup ButtonPressed
    PushButton 145, 90, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
  Text 10, 115, 125, 10, "Current or Upcoming Renewal Month:"
  EditBox 140, 110, 15, 15, Edit2
  EditBox 160, 110, 15, 15, Edit3
  Text 180, 115, 30, 10, "(MM YY)"
  ButtonGroup ButtonPressed
    PushButton 95, 135, 155, 15, "Finish Saving Verbal Attestation Information", finish_btn
EndDialog

MAXIS_case_number = "1529051"
MAXIS_footer_month = "04"
MAXIS_footer_year = "21"

Call MFIP_cert_length_details(verbal_attestation, attestation_verif_array)































'
'
'
' MY_STANDARD_ARRAY = Array("Chris", "Casey", "Aurelia", "Ronin")
' ' all_the_people_in_my_house = all_the_people_in_my_house & "Casey" & "~"
' all_the_people_in_my_house = "Chris~Casey~Aurelia~Ronin~"
' ' all_the_people_in_my_house = left(all_the_people_in_my_house, len(all_the_people_in_my_house)-1)
' MY_STANDARD_ARRAY = Split(all_the_people_in_my_house, "~")
' 	' Dim CLIENT_ARRAY()
' 	' ReDIm CLIENT_ARRAY(0)
' 	'
' 	' Dim CLIENT_ARRAY_WITH_MORE()
' 	' Dim CLIENT_ARRAY_WITH_MORE(0)
' Const ref_numb_const 	= 0
' Const first_name_const 	= 1
' Const last_name_const	= 2
' Const clt_dob_const 	= 3
' Const clt_ssn_last_four_const 	= 4
'
' Dim ALL_CLT_INFO_ARRAY()
' ReDim ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, 0)
'
' the_incrementer = 0
' Do
' 	EmReadscreen MEMB_first_name, 15, 6, 65
' 	EMReadScreen MEMB_last_name, 25, 6, 35
' 	EMReadScreen ref_numb
' 	EMReadScreen dob_mo
' 	EMReadScreen dob_day
' 	EMReadScreen dob_yr
' 	EMReadScreen ssn_last_four
' 	' client_string = client_string & MEMB_first_name & " " & MEMB_last_name & "~"
' 		' ReDim Preserve CLIENT_ARRAY(the_incrementer)
' 		' ReDim Preserve CLIENT_ARRAY_WITH_MORE(the_incrementer)
' 		' CLIENT_ARRAY(the_incrementer) = MEMB_first_name & " " & MEMB_last_name
' 		' CLIENT_ARRAY_WITH_MORE(the_incrementer) = "MEMB " & ref_numb & " - " &  CLIENT_ARRAY(the_incrementer) & " DOB: " & dob_mo & "/" & dob_day & "/" & dob_yr & " SSN: xxx-xx-" & ssn_last_four
' 	ReDim Preserve ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_incrementer)
'
' 	ALL_CLT_INFO_ARRAY(0, the_incrementer) = ref_numb
' 	ALL_CLT_INFO_ARRAY(first_name_const, the_incrementer) = MEMB_first_name
' 	ALL_CLT_INFO_ARRAY(last_name_const, the_incrementer) = MEMB_last_name
' 	ALL_CLT_INFO_ARRAY(clt_dob_const, the_incrementer) = dob_mo & "/" & dob_day & "/" & dob_yr
' 	ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_incrementer) = ssn_last_four
'
' 	the_incrementer = the_incrementer + 1
' 	transmit
' 	EmReadscreen memb_check, 7, 24, 2
' Loop until memb_check = "ENTER A"
' ' CLIENT_ARRAY = split(client_string, "~")
'
' MsgBox Join(MY_STANDARD_ARRAY, ", ")
' For each person in MY_STANDARD_ARRAY
' 	MsgBOx person
' Next
' For the_pers = 0 to UBound(MY_STANDARD_ARRAY)
' 	MsgBOx the_pers
' 	MsgBox MY_STANDARD_ARRAY(the_pers)
' Next
' the_pers = 0
' Do
' 	MsgBox MY_STANDARD_ARRAY(the_pers)
' 	the_pers = the_pers + 1
' Loop until the_pers = UBound(MY_STANDARD_ARRAY)
'
'
'
'
'
'
'
'
'
'
'
' For the_pers = 0 to UBound(ALL_CLT_INFO_ARRAY, 2)				'YOU ALWAYS INCREMENT THE 2nd Parameter of the ARRAY because that is the one that has the different information'
' 	MsgBox "MEMB " & ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers)
'
' 	' last name, first name - dob
' 	MsgBox ALL_CLT_INFO_ARRAY(last_name_const, the_pers) & ", " & ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " - DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers)
'
' 	' dob for MEMB XX - SSN: xxx-xx-____
' 	MsgBox "DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers) & " for MEMB " & ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers) & "- SSN: xxx-xx-" & ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_pers)
' Next
'
'
'
'
'
' function pause_at_certificate_of_understanding()
'     region_known = FALSE        'setting this to start
'     Do
'         EMReadScreen check_for_cert_of_understanding, 28, 2, 28
'         If check_for_cert_of_understanding = "Certificate Of Understanding" Then
'             'go to training region because that is where this thing happens
'             attn            'getting to the primary menu
'             Do
'                 EMReadScreen MAI_check, 3, 1, 33
'                 If MAI_check <> "MAI" then EMWaitReady 1, 1
'             Loop until MAI_check = "MAI"
'
'             If region_known = FALSE Then                        'We only want to look for the region one time - otherwise it would always be Training
'                 region_known = TRUE
'                 EMReadScreen production_status, 7, 6, 15        'looking to see which session was opened
'                 EMReadScreen inquiry_status, 7, 7, 15
'                 EMReadScreen training_status, 7, 8, 15
'                 If production_status = "RUNNING" Then           'Setting a boolean to know which one was opened originally so we can go back to it.
'                     use_prod = TRUE
'                     EMWriteScreen "C", 6, 2     'here we close because otherwise the agreement stays up
'                     transmit
'                 ElseIf inquiry_status = "RUNNING" Then
'                     use_inq = TRUE
'                     EMWriteScreen "C", 7, 2     'here we close because otherwise the agreement stays up
'                     transmit
'                 ElseIf training_status = "RUNNING" Then
'                     use_trn = TRUE
'                 End If
'             End If
'
'             EMWriteScreen "3", 2, 15                        'actually going into training region'
'             transmit
'
'             'Now we stop the script with a dialog so that the user can still interact with MAXIS
'             Dialog1 = ""
'             BeginDialog Dialog1, 0, 0, 211, 155, "MAXIS Certificate of Understanding"
'               ButtonGroup ButtonPressed
'                 OkButton 155, 135, 50, 15
'               Text 5, 5, 135, 15, "It appears it is time for you to review your MAXIS agreement to maintain access."
'               Text 5, 25, 125, 25, "This annual agreement details of using this system in line with privacy and confidentiality requirements."
'               Text 5, 60, 200, 10, "*** YOU MUST READ AND REVIEW THIS INFORMATION ***"
'               GroupBox 5, 75, 200, 55, "Instructions"
'               Text 15, 90, 175, 35, "Leave this dialog up and read the MAXIS screen currently displayed. Enter your agreement selection. Once this is completed, press 'OK' on this dialog and the script will continue. "
'             EndDialog
'
'             Dialog Dialog1                                  'showing the dialog here
'             cancel_without_confirmation
'             'If ButtonPressed = 0 Then stopscript
'         End If
'     Loop until check_for_cert_of_understanding <> "Certificate Of Understanding"    'we keep showing the dialog until this is done
'     If region_known = TRUE Then
'         'Now we are going back to the region we started in.
'         attn
'         Do
'             EMReadScreen MAI_check, 3, 1, 33
'             If MAI_check <> "MAI" then EMWaitReady 1, 1
'         Loop until MAI_check = "MAI"
'         EMWriteScreen "C", 8, 2
'         transmit
'
'         If use_prod = TRUE Then EMWriteScreen "1", 2, 15
'         If use_inq = TRUE Then EMWriteScreen "2", 2, 15
'         If use_trn = TRUE Then EMWriteScreen "3", 2, 15
'         transmit
'     End If
' end function
' EMConnect ""
'
' Call pause_at_certificate_of_understanding
' MsgBox "Moving On"
'
' ' employer_check = MsgBox("Do you have income verification for this job? Employer name: " & "FAMILY DOLLAR", vbYesNo + vbQuestion, "Select Income Panel")
' '
' ' employer_ended_msg = MsgBox("This job has an income end date." & vbNewLine & vbNewLine & "The employer name: FAMILY DOLLAR" & vbNewLine & "End Date: 12/31/19" & vbNewLine & vbNewLine & "The script can update this job with information provided BUT it will remove the 'End Date' field on JOBS." & vbNewLine & vbNewLine & "Would you like to continue with the update of this job?", vbquestion + vbOkCancel, "Income Panel Ended - Cannot Update")
' ' 'Initial Dialog which requests a file path for the excel file
' ' Dialog1 = ""
' ' BeginDialog Dialog1, 0, 0, 361, 105, "On Demand Recertifications"
' '   EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
' '   ButtonGroup ButtonPressed
' '     PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
' '   EditBox 75, 85, 140, 15, worker_signature
' '   ButtonGroup ButtonPressed
' '     OkButton 250, 85, 50, 15
' '     CancelButton 305, 85, 50, 15
' '   Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
' '   Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
' '   Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
' '   Text 10, 90, 60, 10, "Worker Signature"
' ' EndDialog
'
'
' 'Confirmation Diaglog will require worker to afirm the appointment notices/NOMIs should actually be sent
'
' 'END DIALOGS ===============================================================================================================
'
' 'SCRIPT ====================================================================================================================
' 'Connects to BlueZone
' EMConnect ""
'
' Call MAXIS_case_number_finder(MAXIS_case_number)
'
' Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", memb_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("STAT", "JOBS", jobs_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", curr_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("ELIG", "FS  ", fs_priv)
'
' call script_end_procedure("DONE")
'
' Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
' Call script_end_procedure("Case Information" & vbNewLine & vbNewLine & "Case Active - " & case_active & vbNewLine & "Case Pending - " & case_pending & vbNewLine & "Family Cash - " & family_cash_case & vbNewLine &_
'        "MFIP - " & mfip_case & vbNewLine & "DWP - " & dwp_case & vbNewLine & "Adult Cash - " & adult_cash_case & vbNewLine & "GA - " & ga_case & vbNewLine & "MSA - " & msa_case & vbNewLine & "GRH - " & grh_case & vbNewLine &_
'        "SNAP - " & snap_case & vbNewLine & "MA - " & ma_case & vbNewLine & "MSP - " & msp_case & vbNewLine & "CASH Pend - " & unknown_cash_pending)
' 'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
' 'Show initial dialog
' Do
' 	Dialog Dialog1
' 	If ButtonPressed = cancel then stopscript
' 	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
' Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""
'
' 'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
' call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)
'
' 'Set objWorkSheet = objWorkbook.Worksheet
' For Each objWorkSheet In objWorkbook.Worksheets
' 	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
' Next
'
' 'Dialog to select worksheet
' 'DIALOG is defined here so that the dropdown can be populated with the above code
' Dialog1 = ""
' BeginDialog Dialog1, 0, 0, 151, 75, "On Demand Recertifications"
'   DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
'   ButtonGroup ButtonPressed
'     OkButton 40, 55, 50, 15
'     CancelButton 95, 55, 50, 15
'   Text 5, 10, 130, 20, "Select the correct worksheet to run for recertification interview notifications:"
' EndDialog
'
' 'Shows the dialog to select the correct worksheet
' Do
'     Dialog Dialog1
'     If ButtonPressed = cancel then stopscript
' Loop until scenario_dropdown <> "Select One..."
'
' objExcel.worksheets(scenario_dropdown).Activate
'
' excel_row = 2
' leave_loop = FALSE
' Do
'     Call back_to_SELF
'     MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).value)
'
'     If MAXIS_case_number <> "" Then
'         Call navigate_to_MAXIS_screen("STAT", "REVW")
'
'         EMReadScreen cash_revw_status, 1, 7, 40
'         EMReadScreen snap_revw_status, 1, 7, 60
'         EMReadScreen hc_revw_status, 1, 7, 73
'
'         If cash_revw_status = "U" Then leave_loop = TRUE
'         If snap_revw_status = "U" Then leave_loop = TRUE
'         If hc_revw_status = "U" Then leave_loop = TRUE
'         If cash_revw_status = "A" Then leave_loop = TRUE
'         If snap_revw_status = "A" Then leave_loop = TRUE
'         If hc_revw_status = "A" Then leave_loop = TRUE
'
'     Else
'         leave_loop = TRUE
'     End If
'     MAXIS_case_number = ""
'     excel_row = excel_row + 1
'
' Loop until leave_loop = TRUE

Call script_end_procedure_with_error_report("The End")
