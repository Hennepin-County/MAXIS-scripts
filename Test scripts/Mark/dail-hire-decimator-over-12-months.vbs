'STATS GATHERING=============================================================================================================
name_of_script = "BULK - DAIL HIRE DECIMATOR OVER 12 MONTHS.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the actual manual time based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomination applicable to your script.
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
CALL changelog_update("04/17/2024", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function create_array_of_all_active_x_numbers_in_county_with_restart(array_name, two_digit_county_code, restart_status, restart_worker_number)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	Call navigate_to_MAXIS_screen("REPT", "USER")
	PF5 'Hitting PF5 to force sorting, which allows directly selecting a county
	Call write_value_and_transmit(county_code, 21, 6)  	'Inserting county

	MAXIS_row = 7  'Declaring the MAXIS row
	array_name = ""    'Blanking out array_name in case this has been used already in the script

    Found_restart_worker = False    'defaulting to false. Will become true when the X number is found.
	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
            If restart_status = True then
                If trim(UCase(worker_ID)) = trim(UCase(restart_worker_number)) then
                    Found_restart_worker = True
                End if
                If Found_restart_worker = True then array_name = trim(array_name & " " & worker_ID)				'writing to variable
            Else
                array_name = trim(array_name & " " & worker_ID)				'writing to variable
            End if
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then
			PF8			'getting to next screen
			MAXIS_row = 7	're-declaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank

    array_name = split(array_name)
End function

'THE SCRIPT==================================================================================================================
EMConnect ""

'String for HIRE messages that are for cases that cannot be processed - i.e. privileged case
case_numbers_to_skip = "*"
Call Check_for_MAXIS(False)

'Setting script to initially evaluate all x numbers
all_workers_check = 1

this_month = CM_mo & " " & CM_yr
this_month_date = CM_mo & "/01/" & CM_yr
this_month_date = DateAdd("m", 1, this_month_date)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "DAIL Decimator - HIRE Messages over 12 Months Old"
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
  Text 20, 20, 235, 20, "This script should be used to remove HIRE messages that are over 12 months old from current month."
  CheckBox 10, 50, 165, 10, "Check here to process for all workers (default).", all_workers_check
  Text 10, 65, 170, 10, "For restart only, enter the x number to restart from:"
  EditBox 180, 60, 50, 15, restart_worker_number
  ButtonGroup ButtonPressed
    OkButton 155, 90, 50, 15
    CancelButton 210, 90, 50, 15
EndDialog

Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If trim(restart_worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(restart_worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'determining if this is a restart or not in function below when gathering the x numbers.
If trim(restart_worker_number) = "" then
    restart_status = False
Else
	restart_status = True
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_worker_number)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "Deleted DAILs"
ObjExcel.ActiveSheet.Name = "Deleted DAILS"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Value = "FULL DAIL MESSAGE"
objExcel.Cells(1, 7).Value = "ACTION TAKEN"


FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

'This for...next contains each worker indicated above
For each worker in worker_array
    MAXIS_case_number = ""
    back_to_SELF

    'Navigate to DAIL/PICK and select 'INFO' to find HIRE messages
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMWriteScreen "_", 7, 39
    EMWriteScreen "X", 13, 39
    transmit
    
	Call write_value_and_transmit(worker, 21, 6)
	transmit  'transmits past 'not your dail message
    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""
            dail_month = ""
            MAXIS_case_number = ""
			dail_months_old = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
			ELSEIF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
			End if

            dail_row = 6  'resetting the DAIL row

            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
            MAXIS_case_number = trim(MAXIS_case_number)

            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)

            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

			If dail_type = "HIRE" Then

				'Determine if HIRE message is over 12 months old
				dail_month_date = replace(dail_month, " ", "/01/20")
				dail_month_date = dateadd("m", 1, dail_month_date)
				dail_months_old = DateDiff("m", dail_month_date, this_month_date)

				If dail_months_old > 12 Then
					If Instr(case_numbers_to_skip, MAXIS_case_number) = 0 then
						If Instr(dail_msg, "NDNH") or Instr(dail_msg, "JOB DETAILS FOR  ") Then
							
							'Update spreadsheet with DAIL message details
							objExcel.Cells(excel_row, 1).Value = worker
							objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
							objExcel.Cells(excel_row, 3).Value = dail_type
							objExcel.Cells(excel_row, 4).Value = dail_month
							objExcel.Cells(excel_row, 5).Value = dail_msg
							

							'Script will need to read full HIRE message to determine details for INFC

							'Reset variables
							date_hired = ""
							HIRE_employer_name = ""
							priv_case_escape = 1
							hire_match = ""
							full_dail_msg_case_number_only  = ""
							full_dail_msg = ""
                            full_dail_date_hired = ""
                            full_dail_state = ""
							full_dail_msg_line_1 = ""
							full_dail_msg_line_2 = ""
							full_dail_msg_line_3 = ""
							full_dail_msg_line_4 = ""

							'Enters “X” on DAIL message to open full message. 
							Call write_value_and_transmit("X", dail_row, 3)

							'Read entire DAIL message
							EMReadScreen full_dail_msg_line_1, 60, 9, 5
							full_dail_msg_line_1 = trim(full_dail_msg_line_1)
							EMReadScreen full_dail_msg_line_2, 60, 10, 5
							full_dail_msg_line_2 = trim(full_dail_msg_line_2)
							EMReadScreen full_dail_msg_line_3, 60, 11, 5
							full_dail_msg_line_3 = trim(full_dail_msg_line_3)
							EMReadScreen full_dail_msg_line_4, 60, 12, 5
							full_dail_msg_line_4 = trim(full_dail_msg_line_4)

							full_dail_msg = trim(full_dail_msg_case_number & " " & full_dail_msg_case_name & " " & full_dail_msg_line_1 & " " & full_dail_msg_line_2 & " " & full_dail_msg_line_3 & " " & full_dail_msg_line_4)

							'Write full dail message to excel sheet
							objExcel.Cells(excel_row, 6).Value = full_dail_msg

							'Reads MAXIS case number for use in clearing INFC
							EMReadScreen full_dail_msg_case_number_only, 12, 6, 57
                            full_dail_msg_case_number_only = trim(full_dail_msg_case_number_only)

							'Read the NDNH message to find the date hired and convert to MM/DD/YY format
							row = 1
							col = 1
							EMSearch "DATE HIRED   :", row, col
							EMReadScreen full_dail_date_hired, 10, row, col + 15
							If full_dail_date_hired = "  -  -  EM" OR full_dail_date_hired = "UNKNOWN  E" then
								script_end_procedure("Date hired is EM or unknown. Script will now end.")
							End If
							full_dail_date_hired = trim(full_dail_date_hired)
							full_dail_date_hired = Left(full_dail_date_hired, 6) & Right(full_dail_date_hired, 2)

							If Instr(dail_msg, "NDNH") Then
								'Read the state of employment
								row = 1
								col = 1
								EMSearch "NDNH MEMB", row, col
								EMReadScreen full_dail_state, 2, row, col + 17
								full_dail_state = trim(full_dail_state)

							ElseIf Instr(dail_msg, "JOB DETAILS FOR  ") Then
								'Read the state of employment
								row = 1
								col = 1
								EMReadScreen dail_msg_line_1, 74, 9, 5
								dail_msg_line_1 = trim(dail_msg_line_1)
								full_dail_state_array = split(dail_msg_line_1, " ")
								full_dail_state = full_dail_state_array(2)
							End If

							'Identify where ' Employer:' text is so that script can account for slight changes in location in MAXIS
							'Read NDNH message employer
							row = 1
							col = 1
							EMSearch "EMPLOYER: ", row, col
							EMReadScreen full_dail_employer_full_name, 20, row, col + 10
							full_dail_employer_full_name = trim(full_dail_employer_full_name)
							
							'Transmit back to DAIL message
							transmit

							'Navigate to INFC, includes handling to return to DAIL and skip if case is privileged
							Call write_value_and_transmit("I", dail_row, 3)
							EmReadScreen infc_screen_check, 4, 2, 45
							If infc_screen_check <> "INFC" Then
								EmReadScreen self_screen_check, 4, 2, 50
								If self_screen_check = "SELF" Then
									EMReadScreen privileged_check, 22, 24, 2
									If privileged_check = "YOU ARE NOT PRIVILEGED" Then
										EMWriteScreen "DAIL", 16, 43
										EMWriteScreen "________", 18, 43
										EMWriteScreen priv_case_escape, 18, 43
										EMWriteScreen cm_mo, 20, 43
										EMWriteScreen CM_yr, 20, 46
										EMWriteScreen "DAIL", 21, 70
										transmit
										EMReadScreen invalid_case_check, 12, 24, 2
										If invalid_case_check = "INVALID CASE" Then
											Do
												priv_case_escape = priv_case_escape + 1
												EMWriteScreen priv_case_escape, 18, 43
												transmit
												EMReadScreen privileged_check, 22, 24, 2
												If privileged_check <> "INVALID CASE" Then Exit Do
											Loop
										End If
									End If
								End If

								'Get back to HIRE messages for current X number
								EMWriteScreen worker, 21, 6
								transmit
								EMWriteScreen "X", 4, 12
								transmit
								EMWriteScreen "_", 7, 39
								EMWriteScreen "X", 13, 39
								transmit

								'Skip this case moving forward
								case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
								objExcel.Cells(excel_row, 7).Value = "Likely privileged - message skipped."

							Else

								EMReadScreen SSN_present_check, 9, 3, 63
								If SSN_present_check = "_________" Then 
									'Checks if SSN carried forward, if not, it will skip the case moving forward
									objExcel.Cells(excel_row, 7).Value = "Message NOT deleted. SSN is missing so unable to navigate to INFC/HIRE."
									case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 

									'PF3 back to DAIL
									PF3

								Else

									'Navigate to HIRE interface
									Call write_value_and_transmit("HIRE", 20, 71)

									'Handling to ensure script navigated to INFC/HIRE, if not script will end
									EMReadScreen infc_hire_check, 8, 2, 50
									If InStr(infc_hire_check, "HIRE") = 0 Then MsgBox script_end_procedure("Script is unable to navigate to INFC/HIRE. Script will now end.")

									'checking for IRS non-disclosure agreement.
									EMReadScreen agreement_check, 9, 2, 24
									IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

									'Navigate through the interface panel to find the matching employer
									row = 9
									DO
										EMReadScreen infc_case_number, 8, row, 5
										infc_case_number = trim(infc_case_number)
										IF infc_case_number = full_dail_msg_case_number_only THEN
											EMReadScreen infc_employer, 20, row, 36
											infc_employer = trim(infc_employer)
											IF trim(infc_employer) = "" THEN script_end_procedure("An employer match could not be found. The script will now end.")
											IF infc_employer = full_dail_employer_full_name THEN
												EMReadScreen known_by_agency, 1, row, 61
												IF known_by_agency = " " THEN
													EmReadscreen infc_hire_date, 8, row, 20
													EmReadscreen infc_hire_state, 2, row, 31
													infc_hire_state = trim(infc_hire_state)
													If infc_hire_state = "" Then
														If infc_hire_date = full_dail_date_hired Then
															hire_match = TRUE
															match_row = row
															EXIT DO
														End IF
													ElseIf infc_hire_state <> "" Then
														If infc_hire_state = full_dail_state AND infc_hire_date = full_dail_date_hired Then
															hire_match = TRUE
															match_row = row
															EXIT DO
														End If
													End If
												END IF
											END IF
										END IF
										row = row + 1
										IF row = 19 THEN
											PF8
											EmReadscreen end_of_list, 9, 24, 14
											If end_of_list = "LAST PAGE" Then Exit Do
											row = 9
										END IF
									LOOP UNTIL infc_case_number = ""

									IF hire_match <> TRUE THEN 
										'Script failed to clear INFC match, will skip case number moving forward
										objExcel.Cells(excel_row, 7).Value = "Message NOT deleted. No match found in INFC/HIRE."
										case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
									ElseIf hire_match = TRUE Then
										'entering the INFC/HIRE match '
										Call write_value_and_transmit("U", match_row, 3)
										EMReadscreen panel_check, 4, 2, 49
										IF panel_check <> "NHMD" THEN script_end_procedure("Script unable to enter to clear the match. Script will now end")
										EMWriteScreen "Y", 16, 54
										'Agency action must be blank
										TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
										TRANSMIT 'this confirms the cleared status'
										PF3
										EMReadscreen cleared_confirmation, 1, match_row, 61
										IF cleared_confirmation = " " THEN 
											'The total DAILs decreased by 1, message deleted successfully
											objExcel.Cells(excel_row, 7).Value = "Message NOT deleted. No match found in INFC/HIRE."
											case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
										ElseIf cleared_confirmation <> " " THEN 
											'The total DAILs decreased by 1, message deleted successfully
											dail_row = dail_row - 1
											deleted_dails = deleted_dails + 1
											objExcel.Cells(excel_row, 7).Value = "INFC message successfully cleared."
										End If
									End If

									PF3' this takes us back to DAIL/DAIL

									EMReadScreen dail_panel_check, 8, 2, 46
									If InStr(dail_panel_check, "DAIL") = 0 Then 
										PF3
										EMReadScreen dail_panel_check, 8, 2, 46
										If InStr(dail_panel_check, "DAIL") = 0 Then 
											MsgBox "Script unable to return to DAIL"
										End IF
									End If

									EMReadScreen infc_clear_error, 40, 24, 2
									EMReadScreen no_ssn_match_error, 15, 24, 5
									infc_clear_error = trim(infc_clear_error)
									EmReadScreen dail_empty_check, 10, 3, 67
									dail_empty_check = trim(dail_empty_check)
									
									If Instr(infc_clear_error, "THIS IS NOT YOUR DAIL REPORT") and dail_empty_check = "" Then
										'Handling for instances where the DAIL is blank after removing a NDNH message
										PF5

										'Get back to HIRE messages for current X number
										EMWriteScreen worker, 21, 6
										transmit
										EMWriteScreen "X", 4, 12
										transmit
										EMWriteScreen "_", 7, 39
										EMWriteScreen "X", 13, 39
										transmit
									End If

								End If
							End If

							'Increase the excel row
							excel_row = excel_row + 1

							stats_counter = stats_counter + 1   'I increment thee

						ElseIf Instr(dail_msg, "SDNH") or Instr(dail_msg, "NEW JOB DETAILS FOR SSN") Then

							'Update spreadsheet with DAIL message details
							objExcel.Cells(excel_row, 1).Value = worker
							objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
							objExcel.Cells(excel_row, 3).Value = dail_type
							objExcel.Cells(excel_row, 4).Value = dail_month
							objExcel.Cells(excel_row, 5).Value = dail_msg

							'Resetting variables so they do not carry forward
							last_dail_check = ""
							other_worker_error = ""
							total_dail_msg_count_before = ""
							total_dail_msg_count_after = ""
							all_done = ""
							final_dail_error = ""
							full_dail_msg_line_1 = ""
							full_dail_msg_line_2 = ""
							full_dail_msg_line_3 = ""
							full_dail_msg_line_4 = ""
							full_dail_msg = ""
							
							'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
							EMReadScreen last_dail_check, 12, 3, 67
							last_dail_check = trim(last_dail_check)

							'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
							last_dail_check = split(last_dail_check, " ")

							If last_dail_check(0) = last_dail_check(2) then 
								'The script is about to delete the LAST message in the DAIL so script will exit do loop after deletion, also works if it is about to delete the ONLY message in the DAIL
								all_done = true
							End If

							'Open full Dail message
							Call write_value_and_transmit("X", dail_row, 3)

							'Capture full dail message details
							EMReadScreen full_dail_msg_line_1, 60, 9, 5
							full_dail_msg_line_1 = trim(full_dail_msg_line_1)
							EMReadScreen full_dail_msg_line_2, 60, 10, 5
							full_dail_msg_line_2 = trim(full_dail_msg_line_2)
							EMReadScreen full_dail_msg_line_3, 60, 11, 5
							full_dail_msg_line_3 = trim(full_dail_msg_line_3)
							EMReadScreen full_dail_msg_line_4, 60, 12, 5
							full_dail_msg_line_4 = trim(full_dail_msg_line_4)

							full_dail_msg = trim(full_dail_msg_case_number & " " & full_dail_msg_case_name & " " & full_dail_msg_line_1 & " " & full_dail_msg_line_2 & " " & full_dail_msg_line_3 & " " & full_dail_msg_line_4)

							'Write full dail message to excel sheet
							objExcel.Cells(excel_row, 6).Value = full_dail_msg

							'Transmit back to DAIL
							transmit

							'Delete the message
							Call write_value_and_transmit("D", dail_row, 3)

							'Handling for deleting message under someone else's x number
							EMReadScreen other_worker_error, 25, 24, 2
							other_worker_error = trim(other_worker_error)

							If other_worker_error = "ALL MESSAGES WERE DELETED" Then
								'Script deleted the final message in the DAIL
								dail_row = dail_row - 1
								deleted_dails = deleted_dails + 1
								objExcel.Cells(excel_row, 7).Value = "Message deleted."

								'Exit do loop as all messages are deleted
								all_done = true

							ElseIf other_worker_error = "" Then
								'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

								'Handling to check if message actually deleted
								total_dail_msg_count_before = last_dail_check(2) * 1
								EMReadScreen total_dail_msg_count_after, 12, 3, 67

								total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
								total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

								If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
									'The total DAILs decreased by 1, message deleted successfully
									dail_row = dail_row - 1
									deleted_dails = deleted_dails + 1
									objExcel.Cells(excel_row, 7).Value = "Message deleted."
								Else
									'The total DAILs did not decrease by 1, something went wrong
									objExcel.Cells(excel_row, 7).Value = "Likely privileged or some other issue so unable to delete - message skipped."
									'Skip this case moving forward
									case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 

								End If

							ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
								
								'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
								transmit

								'Reads the total number of DAILS after deleting to determine if it decreased by 1
								EMReadScreen total_dail_msg_count_after, 12, 3, 67

								'Checks if final DAIL message deleted
								EMReadScreen final_dail_error, 25, 24, 2

								If final_dail_error = "ALL MESSAGES WERE DELETED" Then
									'All DAIL messages deleted so indicates deletion a success
									dail_row = dail_row - 1
									deleted_dails = deleted_dails + 1
									objExcel.Cells(excel_row, 7).Value = "Message deleted."
									'No more DAIL messages so exit do loop
									all_done = True
								ElseIf trim(final_dail_error) = "" Then
									'Handling to check if message actually deleted
									total_dail_msg_count_before = last_dail_check(2) * 1

									total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
									total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

									If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
										'The total DAILs decreased by 1, message deleted successfully
										dail_row = dail_row - 1
										deleted_dails = deleted_dails + 1
										objExcel.Cells(excel_row, 7).Value = "Message deleted."
									Else
										'The total DAILs did not decrease by 1, something went wrong
										objExcel.Cells(excel_row, 7).Value = "Likely privileged or some other issue so unable to delete - message skipped."
										'Skip this case moving forward
										case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
									End If

								Else
									'The total DAILs did not decrease by 1, something went wrong
									objExcel.Cells(excel_row, 7).Value = "Likely privileged or some other issue so unable to delete - message skipped."
									'Skip this case moving forward
									case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
								End if
								
							Else
								'The total DAILs did not decrease by 1, something went wrong
								objExcel.Cells(excel_row, 7).Value = "Likely privileged or some other issue so unable to delete - message skipped."
								'Skip this case moving forward
								case_numbers_to_skip = case_numbers_to_skip & MAXIS_case_number & "*" 
							End If
				
							excel_row = excel_row + 1
							stats_counter = stats_counter + 1   'I increment thee

						Else
							'Handling just in case an unusual HIRE message comes up
							
							objExcel.Cells(excel_row, 1).Value = worker
							objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
							objExcel.Cells(excel_row, 3).Value = dail_type
							objExcel.Cells(excel_row, 4).Value = dail_month
							objExcel.Cells(excel_row, 5).Value = dail_msg
							objExcel.Cells(excel_row, 7).Value = "Not a NDNH or SDNH HIRE message."

							excel_row = excel_row + 1

						End If
					End If
				End If
			ElseIf dail_type = "TIKL" or dail_type = "CSES" or dail_type = "PEPR" Then
				'If for some reason, ALL DAILs are showing again then it will reset to HIRE only
				EMWriteScreen worker, 21, 6
				transmit
				EMWriteScreen "X", 4, 12
				transmit
				EMWriteScreen "_", 7, 39
				EMWriteScreen "X", 13, 39
				transmit
				dail_row = dail_row - 1

			End If

			'Go to next dail_row
			dail_row = dail_row + 1

            'checking for the last DAIL message - If it's the last message, which can be blank OR _ then the script will exit the do. 
			EMReadScreen next_dail_check, 7, dail_row, 3
			If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                PF8
                EMReadScreen next_dail_check, 7, dail_row, 3
			    If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    last_case = true
				    exit do
                End if 
			End if
		LOOP
		IF last_case = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 8).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 8).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 8).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 8).Value = "Script run time (in seconds):"
objExcel.Cells(6, 8).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 8).Value = "Number of messages reviewed/DAIL messages remaining:"
objExcel.Cells(8, 8).Value = "False count/duplicate DAIL Messages not counted:"
objExcel.Columns(8).Font.Bold = true
objExcel.Cells(2, 9).Value = deleted_dails
objExcel.Cells(3, 9).Value = STATS_manualtime
objExcel.Cells(4, 9).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 9).Value = timer - start_time
objExcel.Cells(6, 9).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 9).Value = STATS_counter
objExcel.Cells(8, 9).Value = false_count

'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " " & "HIRE Messages over 12 months old" & " " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("Success! HIRE messages over 12 months old have been cleared.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/01/2024
'--Tab orders reviewed & confirmed----------------------------------------------05/01/2024
'--Mandatory fields all present & Reviewed--------------------------------------05/01/2024
'--All variables in dialog match mandatory fields-------------------------------05/01/2024
'Review dialog names for content and content fit in dialog----------------------05/01/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/01/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------05/01/2024
'--PRIV Case handling reviewed -------------------------------------------------05/01/2024
'--Out-of-County handling reviewed----------------------------------------------05/01/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/01/2024
'--BULK - review output of statistics and run time/count (if applicable)--------05/01/2024
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/01/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------05/01/2024
'--Denomination reviewed -------------------------------------------------------05/01/2024
'--Script name reviewed---------------------------------------------------------05/01/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/01/2024

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/01/2024
'--comment Code-----------------------------------------------------------------05/01/2024
'--Update Changelog for release/update------------------------------------------05/01/2024
'--Remove testing message boxes-------------------------------------------------05/01/2024
'--Remove testing code/unnecessary code-----------------------------------------05/01/2024
'--Review/update SharePoint instructions----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/01/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A

