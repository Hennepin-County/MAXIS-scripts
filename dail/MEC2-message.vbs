'STATS GATHERING=============================================================================================================
name_of_script = "DAIL - MEC2 Message.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("01/15/25", "Initial version.", "Mark Riegel, Hennepin County") 

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone

'Read the dail message
' EmWriteScreen "X", 6, 3
' transmit
'Reads the entire line above the DAIL message (the full case name and case number)
EmReadScreen full_case_name_number, 76, 5, 5
MAXIS_case_number = trim(right(full_case_name_number, 8))
dail_message_member_name = left(full_case_name_number, 4)
Call write_value_and_transmit("X", 6, 3)
EMReadScreen full_message_1, 70, 9, 5
EMReadScreen full_message_2, 70, 10, 5
EMReadScreen full_message_3, 70, 11, 5
EMReadScreen full_message_4, 70, 12, 5
full_message = trim(trim(full_message_1) & " " & trim(full_message_2) & " " & trim(full_message_3) & " " & trim(full_message_4))
full_case_name_number_message = full_case_name_number & full_message
'Transmit back to DAIL
transmit

If instr(full_message, "RSDI END DATE") OR instr(full_message, "SSI REPORTED TO MEC²") OR instr(full_message, "UNEMPLOYMENT INS") OR instr(full_message, "SELF EMPLOYMENT REPORTED TO MEC²") Then
	If instr(full_message, "RSDI END DATE") OR instr(full_message, "SSI REPORTED TO MEC²") OR instr(full_message, "UNEMPLOYMENT INS") THEN
		'Navigate to STAT/UNEA
		Call write_value_and_transmit("S", 6, 3)
		Call write_value_and_transmit("UNEA", 20, 71)
	ElseIf instr(full_message, "SELF EMPLOYMENT REPORTED TO MEC²") Then
		'Navigate to STAT/BUSI
		Call write_value_and_transmit("S", 6, 3)
		Call write_value_and_transmit("BUSI", 20, 71)
	End If

	Dialog1 = "" 'blanking out dialog name
	BeginDialog Dialog1, 0, 0, 311, 150, "DAIL - MEC2 Message"
		ButtonGroup ButtonPressed
		PushButton 5, 90, 65, 15, "HSR Manual", hsr_manual_btn
		PushButton 5, 110, 65, 15, "Script Instructions", script_instructions_btn
		OkButton 205, 130, 50, 15
		CancelButton 255, 130, 50, 15
		Text 5, 5, 55, 10, "DAIL Message"
		Text 5, 20, 295, 35, full_message
		Text 5, 65, 300, 20, "The script has navigated to the respective STAT panel noted in the DAIL message. Please review the change and take any action required by the message. "
		Text 75, 95, 85, 10, "Link to HSR Manual"
		Text 75, 115, 85, 10, "Link to Script Instructions"
	EndDialog

	DO
		Do
			err_msg = ""    'This is the error message handling
			Dialog Dialog1
			cancel_without_confirmation
			If ButtonPressed = hsr_manual_btn Then 
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MEC2.aspx"
				err_msg = "LOOP"
			End If
			If ButtonPressed = script_instructions_btn Then 
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%20MEC2%20Message.docx"
				err_msg = "LOOP"
			End If
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	'End the script.
	script_end_procedure("Please follow the instructions provided in the HSR Manual. The script will now end.")

Else
	'Message can be deleted
	' Msgbox "This MEC2 message is non-actionable and will be deleted. Press 'OK' to delete the message. Press the 'X' to stop the script."

	Dialog1 = "" 'blanking out dialog name
	BeginDialog Dialog1, 0, 0, 311, 115, "DAIL - MEC2 Message"
		ButtonGroup ButtonPressed
		OkButton 205, 95, 50, 15
		CancelButton 255, 95, 50, 15
		Text 5, 5, 55, 10, "DAIL Message"
		Text 5, 20, 300, 35, full_message
		Text 5, 65, 300, 20, "This MEC2 message is non-actionable and will be deleted. Press 'OK' to delete the message. Press 'Cancel' to stop the script."
	EndDialog

	DO
		Dialog Dialog1
		cancel_without_confirmation
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	If ButtonPressed = OK Then
		'Ensure we are still at the DAIL and the same message
		EmReadScreen dail_panel_check, 4, 2, 48
		If dail_panel_check <> "DAIL" Then
			'No longer at the DAIL so script must navigate back to the correct DAIL message
			PF3_count = 0
			Do
				PF3
				EmReadScreen at_dail_check, 4, 2, 48
				If at_dail_check = "DAIL" Then
					Exit Do
				Else
					PF3_count = PF3_count + 1
				End If
				If PF3_count = 4 Then
					unable_to_PF3_to_dail = True
					Exit Do
				End If
			Loop
			
			If unable_to_PF3_to_dail = True Then
				back_to_SELF
				EMReadScreen self_panel_check, 4, 2, 50
				If self_panel_check <> "SELF" Then

					'Script will end if unable to get back to SELF to then get back to DAIL
					script_end_procedure("The script is unable to navigate back to the DAIL. The script will now end.")
				Else

					'Navigate to DAIL
					EMWriteScreen "DAIL", 16, 43
					EMWriteScreen "DAIL", 21, 70
					transmit

					EMReadScreen back_to_dail_check, 8, 1, 72

					If back_to_dail_check <> "FMKDLAM6" Then
						'Script will end if unable to get back to SELF to then get back to DAIL
						script_end_procedure("The script is unable to navigate back to the DAIL. The script will now end.")

					Else 
						'To do - I don't think these are needed but commenting just in case
						' 'Navigate to CASE/CURR to force DAIL to reset and then PF3 back to get back to start of the DAIL
						' Call write_value_and_transmit("H", dail_row, 3)
						' PF3

						'Update DAIL/PICK to MEC2
						Call write_value_and_transmit("X", 4, 12)
						EmWriteScreen "_", 7, 39
						Call write_value_and_transmit("X", 16, 39)

						'Script should now navigate to specific member name, or at least get close
						EMWriteScreen dail_message_member_name, 21, 25
						transmit

						'Script will enter do loop to find match
						'Set dail_row to 6 at start
						dail_row = 6

						Do
							'Ensure we are at the same message before deleting
							EmReadScreen check_full_case_name_number, 76, 5, 5
							Call write_value_and_transmit("X", 6, 3)
							EMReadScreen check_full_message_1, 70, 9, 5
							EMReadScreen check_full_message_2, 70, 10, 5
							EMReadScreen check_full_message_3, 70, 11, 5
							EMReadScreen check_full_message_4, 70, 12, 5
							check_full_case_name_number_message = full_case_name_number & trim(trim(check_full_message_1) & " " & trim(check_full_message_2) & " " & trim(check_full_message_3) & " " & trim(check_full_message_4))
							transmit

							If check_full_case_name_number_message = full_case_name_number_message Then
								transmit
								Exit Do
							Else
								transmit
								dail_row = dail_row + 1

								'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
								EMReadScreen new_case, 8, dail_row, 63
								new_case = trim(new_case)
								IF new_case <> "CASE NBR" THEN 
									'If there is NOT a new case number, the script will top the message
									Call write_value_and_transmit("T", dail_row, 3)
								ELSEIF new_case = "CASE NBR" THEN
									'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
									Call write_value_and_transmit("T", dail_row + 1, 3)
								End if
							End If

						Loop
					End If
				End If

			Else
				'Script got back to DAIL with PF3s
				'Update DAIL/PICK to MEC2
				Call write_value_and_transmit("X", 4, 12)
				EmWriteScreen "_", 7, 39
				Call write_value_and_transmit("X", 16, 39)

				'Script should now navigate to specific member name, or at least get close
				EMWriteScreen dail_message_member_name, 21, 25
				transmit

				'Script will enter do loop to find match
				'Set dail_row to 6 at start
				dail_row = 6

				Do
					'Ensure we are at the same message before deleting
					EmReadScreen check_full_case_name_number, 76, 5, 5
					Call write_value_and_transmit("X", 6, 3)
					EMReadScreen check_full_message_1, 70, 9, 5
					EMReadScreen check_full_message_2, 70, 10, 5
					EMReadScreen check_full_message_3, 70, 11, 5
					EMReadScreen check_full_message_4, 70, 12, 5
					check_full_case_name_number_message = full_case_name_number & trim(trim(check_full_message_1) & " " & trim(check_full_message_2) & " " & trim(check_full_message_3) & " " & trim(check_full_message_4))
					transmit

					If check_full_case_name_number_message = full_case_name_number_message Then
						transmit
						Exit Do
					Else
						transmit
						dail_row = dail_row + 1

						'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
						EMReadScreen new_case, 8, dail_row, 63
						new_case = trim(new_case)
						IF new_case <> "CASE NBR" THEN 
							'If there is NOT a new case number, the script will top the message
							Call write_value_and_transmit("T", dail_row, 3)
						ELSEIF new_case = "CASE NBR" THEN
							'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
							Call write_value_and_transmit("T", dail_row + 1, 3)
						End if
					End If
				Loop
			End If
		End If
	End If
End if

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
