''GATHERING STATS===========================================================================================
name_of_script = "ACTION - DEU MATCH CLEARED CC.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches and CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match when the date is over 45 days.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/14/2017", "Updated to fix claim entering and case note header.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/14/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Fun with dates! --Creating variables for the rolling 12 calendar months
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month -2'
CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
'current month -3'
CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)
'current month -4'
CM_minus_4_mo =  right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 2)
CM_minus_4_yr =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 2)
'current month -5'
CM_minus_5_mo =  right("0" &             DatePart("m",           DateAdd("m", -5, date)            ), 2)
CM_minus_5_yr =  right(                  DatePart("yyyy",        DateAdd("m", -5, date)            ), 2)
'current month -6'
CM_minus_6_mo =  right("0" &             DatePart("m",           DateAdd("m", -6, date)            ), 2)
CM_minus_6_yr =  right(                  DatePart("yyyy",        DateAdd("m", -6, date)            ), 2)
'current month -7'
CM_minus_7_mo =  right("0" &             DatePart("m",           DateAdd("m", -7, date)            ), 2)
CM_minus_7_yr =  right(                  DatePart("yyyy",        DateAdd("m", -7, date)            ), 2)
'current month -8'
CM_minus_8_mo =  right("0" &             DatePart("m",           DateAdd("m", -8, date)            ), 2)
CM_minus_8_yr =  right(                  DatePart("yyyy",        DateAdd("m", -8, date)            ), 2)
'current month -9'
CM_minus_9_mo =  right("0" &             DatePart("m",           DateAdd("m", -9, date)            ), 2)
CM_minus_9_yr =  right(                  DatePart("yyyy",        DateAdd("m", -9, date)            ), 2)
'current month -10'
CM_minus_10_mo =  right("0" &            DatePart("m",           DateAdd("m", -10, date)           ), 2)
CM_minus_10_yr =  right(                 DatePart("yyyy",        DateAdd("m", -10, date)           ), 2)
'current month -11'
CM_minus_11_mo =  right("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'Establishing value of variables for the rolling 12 months
current_month = CM_mo & "/" & CM_yr
current_month_minus_one = CM_minus_1_mo & "/" & CM_minus_1_yr
current_month_minus_two = CM_minus_2_mo & "/" & CM_minus_2_yr
current_month_minus_three = CM_minus_3_mo & "/" & CM_minus_3_yr
current_month_minus_four = CM_minus_4_mo & "/" & CM_minus_4_yr
current_month_minus_five = CM_minus_5_mo & "/" & CM_minus_5_yr
current_month_minus_six = CM_minus_6_mo & "/" & CM_minus_6_yr
current_month_minus_seven = CM_minus_7_mo & "/" & CM_minus_7_yr
current_month_minus_eight = CM_minus_8_mo & "/" & CM_minus_8_yr
current_month_minus_nine = CM_minus_9_mo & "/" & CM_minus_9_yr
current_month_minus_ten = CM_minus_10_mo & "/" & CM_minus_10_yr
current_month_minus_eleven = CM_minus_11_mo & "/" & CM_minus_11_yr

'---------------------------------------------------------------------THE SCRIPT
EMConnect ""

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)
memb_number = "01"
discovery_date = date & ""
Call navigate_to_MAXIS_screen("CASE", "NOTE")
EmWriteScreen "x", 5, 3
Transmit
note_row = 4			'Beginning of the case notes
Do 						'Read each line
	EMReadScreen note_line, 76, note_row, 3
	     	note_line = trim(note_line)
	If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
	message_array = message_array & note_line & vbcr		'putting the lines together
	note_row = note_row + 1
	If note_row = 18 then 									'End of a single page of the case note
		EMReadScreen next_page, 7, note_row, 3
		If next_page = "More: +" Then 						'This indicates there is another page of the case note
			PF8												'goes to the next line and resets the row to read'\
			note_row = 4
		End If
	End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
	    'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	    CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & OP_Date & "Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)
	END IF
END IF

'---------------------------------------------------------------writing the CCOL case note'
	Call navigate_to_MAXIS_screen("CCOL", "CLSM")
	EMWriteScreen Claim_number, 4, 9
	Transmit
	EMReadScreen existing_case_note, 1, 5, 6
	IF existing_case_note = "" THEN
		PF4
	ELSE
		PF9
	END IF
	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH" & "(" & first_name &  ")" & "CLEARED CC-CLAIM ENTERED-----")
	IF IEVS_type = "BEER" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & type_match & ") " & "(" & first_name &  ")" &  "CLEARED CC-CLAIM ENTERED-----")
	CALL write_bullet_and_variable_in_CCOL_NOTE("Discovery date", discovery_date)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Period", IEVS_period)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Active Programs", programs)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Source of income", source_income)
	CALL write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
	CALL write_variable_in_CCOL_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
	IF OP_program_II <> "Select:" then CALL write_variable_in_CCOL_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
	IF OP_program_III <> "Select:" then CALL write_variable_in_CCOL_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
	IF OT_resp_memb <> "" THEN CALL write_bullet_and_variable_in_CCOL_NOTE("Other responsible member(s)", OT_resp_memb)
	IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_NOTE("* Earned Income Disregard Allowed")
	IF programs = "Health Care" or programs = "Medical Assistance" THEN
		Call write_bullet_and_variable_in_CCOL_NOTE("HC responsible members", HC_resp_memb)
		Call write_bullet_and_variable_in_CCOL_NOTE("HC claim number", hc_claim_number)
		Call write_bullet_and_variable_in_CCOL_NOTE("Total federal Health Care amount", Fed_HC_AMT)
		CALL write_variable_in_CCOL_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
	END IF
	CALL write_bullet_and_variable_in_CCOL_NOTE("Fraud referral made", fraud_referral)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Collectible claim", collectible_dropdown)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Reason that claim is collectible or not", collectible_reason)
	CALL write_bullet_and_variable_in_CCOL_NOTE("Income verification received", income_rcvd_date)
	CALL write_bullet_and_variable_in_CCOL_NOTE("MANDATORY-Reason for overpayment", Reason_OP)
	CALL write_variable_in_CCOL_NOTE("----- ----- ----- ----- ----- ----- -----")
	CALL write_variable_in_CCOL_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
	PF3 'exit the case note'
	PF3 'back to dail'
Script_end_procedure("Overpayment case note entered and copied to CCOL.")
