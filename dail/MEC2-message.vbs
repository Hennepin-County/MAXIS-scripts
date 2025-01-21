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

'Reads the entire line above the DAIL message (the full case name and case number)
EmReadScreen full_case_name_number, 76, 5, 5
dail_message_member_name = left(full_case_name_number, 4)
MAXIS_case_number = trim(right(full_case_name_number, 8))
'Opens the message to capture the full text of the message
Call write_value_and_transmit("X", 6, 3)
EMReadScreen full_message_1, 70, 9, 5
EMReadScreen full_message_2, 70, 10, 5
EMReadScreen full_message_3, 70, 11, 5
EMReadScreen full_message_4, 70, 12, 5
full_message = trim(trim(full_message_1) & " " & trim(full_message_2) & " " & trim(full_message_3) & " " & trim(full_message_4))
'Creates variable of full case name and number with full message text so that the message can be found later if needed
full_case_name_number_message = full_case_name_number & full_message
'Transmit back to DAIL
transmit

'If the MEC2 message is one of the following, then it is actionable and cannot be deleted.
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

	'Dialog provides links to HSR manual and script instructions. Worker will need to act on the message depending on the situation
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
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/ALL%20DAIL%20SCRIPTS.docx"
				err_msg = "LOOP"
			End If
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	'End the script.
	script_end_procedure("Please follow the instructions provided in the HSR Manual. The script will now end.")

Else
	'If the MEC2 message is not one of the ones noted above, then it can be deleted as it is non-actionable

	'Dialog informing worker that message will be deleted
	Dialog1 = "" 'blanking out dialog name
	BeginDialog Dialog1, 0, 0, 306, 115, "DAIL - MEC2 Message"
		ButtonGroup ButtonPressed
		OkButton 205, 95, 50, 15
		CancelButton 255, 95, 50, 15
		Text 5, 5, 55, 10, "DAIL Message"
		Text 5, 20, 300, 35, full_message
		Text 5, 65, 300, 20, "This MEC2 message is non-actionable and will be deleted. Press 'OK' to delete the message. Press 'Cancel' to stop the script."
		ButtonGroup ButtonPressed
		PushButton 5, 95, 65, 15, "Script Instructions", script_instructions_btn
	EndDialog

	DO
		Do
			err_msg = ""    'This is the error message handling
			Dialog Dialog1
			cancel_without_confirmation
			If ButtonPressed = script_instructions_btn Then 
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/ALL%20DAIL%20SCRIPTS.docx"
				err_msg = "LOOP"
			End If
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	If ButtonPressed = OK Then
		'Ensure we are still at the DAIL and the same message
		EmReadScreen dail_panel_check, 4, 2, 48
		If dail_panel_check <> "DAIL" Then
			'If we are no longer at the DAIL, then the script must navigate back to the DAIL so it can delete the correct message. It will start with using PF3 to navigate back
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
			
			'If the script didn't make it to the DAIL after 3 PF3s
			If unable_to_PF3_to_dail = True Then
				'Navigate to SELF to then navigate back to the DAIL
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
						'Script will end if unable to navigate to DAIL from SELF
						script_end_procedure("The script is unable to navigate back to the DAIL. The script will now end.")
					Else 
						
						'Update DAIL/PICK to MEC2
						Call write_value_and_transmit("X", 4, 12)
						EmWriteScreen "_", 7, 39
						Call write_value_and_transmit("X", 16, 39)

						'Find previous case number and then case name
						Call write_value_and_transmit(MAXIS_case_number, 20, 38)
						Call write_value_and_transmit(dail_message_member_name, 21, 25)

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
							check_full_case_name_number_message = check_full_case_name_number & trim(trim(check_full_message_1) & " " & trim(check_full_message_2) & " " & trim(check_full_message_3) & " " & trim(check_full_message_4))
							'Exit message and transmit back to DAIL
							transmit	

							If check_full_case_name_number_message = full_case_name_number_message Then
								'Matching message found, it will delete and then the script will end

								'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
								EMReadScreen last_dail_check, 12, 3, 67
								last_dail_check = trim(last_dail_check)
		
								'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
								last_dail_check = split(last_dail_check, " ")
								

								Call write_value_and_transmit("D", dail_row, 3)

								'Handling for deleting message under someone else's x number
								EMReadScreen other_worker_error, 25, 24, 2
								other_worker_error = trim(other_worker_error)
		
								If other_worker_error = "ALL MESSAGES WERE DELETED" Then
									'Script deleted the final message in the DAIL successfully 
									script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
		
								ElseIf other_worker_error = "" Then
									'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion
		
									'Handling to check if message actually deleted
									total_dail_msg_count_before = last_dail_check(2) * 1
									EMReadScreen total_dail_msg_count_after, 12, 3, 67
		
									total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
									total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1
		
									If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
										'The total DAILs decreased by 1, message deleted successfully
										script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
									Else
										'The message deletion failed
										script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
									End If
		
								ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
									
									'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
									transmit
		
									'Reads the total number of DAILS after deleting to determine if it decreased by 1
									EMReadScreen total_dail_msg_count_after, 12, 3, 67
		
									'Checks if final DAIL message deleted
									EMReadScreen final_dail_error, 25, 24, 2
		
									If final_dail_error = "ALL MESSAGES WERE DELETED" Then
										'Script deleted the final message in the DAIL successfully 
										script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
									ElseIf trim(final_dail_error) = "" Then
										'Handling to check if message actually deleted
										total_dail_msg_count_before = last_dail_check(2) * 1
		
										total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
										total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1
		
										If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
											'Script deleted the final message in the DAIL successfully 
											script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
										Else
											'The message deletion failed
											script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
										End If
		
									Else
										'The message deletion failed
										script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
									End if
									
								Else
									'The message deletion failed
									script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
								End If
							Else
								'If message is not a match, it will move to the next DAIL message to check
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

								'Reset dail_row to 6
								dail_row = 6
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

				'Find previous case number and then case name
				Call write_value_and_transmit(MAXIS_case_number, 20, 38)
				Call write_value_and_transmit(dail_message_member_name, 21, 25)

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
					check_full_case_name_number_message = check_full_case_name_number & trim(trim(check_full_message_1) & " " & trim(check_full_message_2) & " " & trim(check_full_message_3) & " " & trim(check_full_message_4))
					'Exit message and transmit back to DAIL
					transmit

					If check_full_case_name_number_message = full_case_name_number_message Then
						'Matching message found, it will delete and then the script will end

						'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
						EMReadScreen last_dail_check, 12, 3, 67
						last_dail_check = trim(last_dail_check)

						'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
						last_dail_check = split(last_dail_check, " ")
						

						Call write_value_and_transmit("D", dail_row, 3)

						'Handling for deleting message under someone else's x number
						EMReadScreen other_worker_error, 25, 24, 2
						other_worker_error = trim(other_worker_error)

						If other_worker_error = "ALL MESSAGES WERE DELETED" Then
							'Script deleted the final message in the DAIL successfully 
							script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")

						ElseIf other_worker_error = "" Then
							'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

							'Handling to check if message actually deleted
							total_dail_msg_count_before = last_dail_check(2) * 1
							EMReadScreen total_dail_msg_count_after, 12, 3, 67

							total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
							total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

							If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
								'The total DAILs decreased by 1, message deleted successfully
								script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
							Else
								'The message deletion failed
								script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
							End If

						ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
							
							'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
							transmit

							'Reads the total number of DAILS after deleting to determine if it decreased by 1
							EMReadScreen total_dail_msg_count_after, 12, 3, 67

							'Checks if final DAIL message deleted
							EMReadScreen final_dail_error, 25, 24, 2

							If final_dail_error = "ALL MESSAGES WERE DELETED" Then
								'Script deleted the final message in the DAIL successfully 
								script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
							ElseIf trim(final_dail_error) = "" Then
								'Handling to check if message actually deleted
								total_dail_msg_count_before = last_dail_check(2) * 1

								total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
								total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

								If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
									'Script deleted the final message in the DAIL successfully 
									script_end_procedure("The script successfully deleted the MEC2 message. The script will now end.")
								Else
									'The message deletion failed
									script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
								End If

							Else
								'The message deletion failed
								script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
							End if
							
						Else
							'The message deletion failed
							script_end_procedure("The script was unable to delete the MEC2 message. Please delete the message manually. The script will now end.")
						End If
					Else
						'If message is not a match, it will move to the next DAIL message to check
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

						'Reset dail_row to 6
						dail_row = 6
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
'--Dialog1 = "" on all dialogs -------------------------------------------------01/21/2025
'--Tab orders reviewed & confirmed----------------------------------------------01/21/2025
'--Mandatory fields all present & Reviewed--------------------------------------01/21/2025
'--All variables in dialog match mandatory fields-------------------------------01/21/2025
'Review dialog names for content and content fit in dialog----------------------01/21/2025
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------01/21/2025
'--Include script category and name somewhere on first dialog-------------------01/21/2025
'--Create a button to reference instructions------------------------------------01/21/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------01/21/2025
'--MAXIS_background_check reviewed (if applicable)------------------------------01/21/2025
'--PRIV Case handling reviewed -------------------------------------------------01/21/2025
'--Out-of-County handling reviewed----------------------------------------------01/21/2025
'--script_end_procedures (w/ or w/o error messaging)----------------------------01/21/2025
'--BULK - review output of statistics and run time/count (if applicable)--------01/21/2025
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------01/21/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------01/21/2025
'--Incrementors reviewed (if necessary)-----------------------------------------01/21/2025
'--Denomination reviewed -------------------------------------------------------01/21/2025
'--Script name reviewed---------------------------------------------------------01/21/2025
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------01/21/2025
'--comment Code-----------------------------------------------------------------01/21/2025
'--Update Changelog for release/update------------------------------------------01/21/2025
'--Remove testing message boxes-------------------------------------------------01/21/2025
'--Remove testing code/unnecessary code-----------------------------------------01/21/2025
'--Review/update SharePoint instructions----------------------------------------01/21/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------01/21/2025
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------01/21/2025
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------01/21/2025
'--Complete misc. documentation (if applicable)---------------------------------01/21/2025
'--Update project team/issue contact (if applicable)----------------------------01/21/2025
