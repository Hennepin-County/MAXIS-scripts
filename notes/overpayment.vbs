'GATHERING STATS===========================================================================================
name_of_script = "NOTES - OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/25/2018", "Updated script to copy case note to CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/01/2018", "Updated script to ensure Reason for OP is entered as it is a mandatory field.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/25/2018", "Updated script to add Fraud and Earned Income handling.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Updated script to add HC handling and the income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/01/2018", "Updated script to write amount in case note in the correct area.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/04/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------FUNCTION'

EMConnect ""

CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
discovery_date = date & ""
back_to_self

	BeginDialog overpayment_dialog, 0, 0, 361, 230, "Overpayment Claim Entered"
	  EditBox 60, 5, 40, 15, MAXIS_case_number
	  EditBox 140, 5, 20, 15, memb_number
	  EditBox 230, 5, 20, 15, OT_resp_memb
	  DropListBox 315, 5, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	  EditBox 60, 25, 40, 15, discovery_date
	  DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
	  EditBox 130, 65, 30, 15, OP_from
	  EditBox 180, 65, 30, 15, OP_to
	  EditBox 245, 65, 35, 15, Claim_number
	  EditBox 305, 65, 45, 15, Claim_amount
	  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
	  EditBox 130, 85, 30, 15, OP_from_II
	  EditBox 180, 85, 30, 15, OP_to_II
	  EditBox 245, 85, 35, 15, Claim_number_II
	  EditBox 305, 85, 45, 15, Claim_amount_II
	  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
	  EditBox 130, 105, 30, 15, OP_from_III
	  EditBox 180, 105, 30, 15, OP_to_III
	  EditBox 245, 105, 35, 15, Claim_number_III
	  EditBox 305, 105, 45, 15, Claim_amount_III
	  EditBox 190, 125, 20, 15, HC_resp_memb
	  EditBox 305, 125, 45, 15, Fed_HC_AMT
	  EditBox 70, 150, 160, 15, income_source
	  CheckBox 235, 155, 120, 10, "Earned Income Disregard Allowed", EI_checkbox
	  EditBox 70, 170, 160, 15, EVF_used
	  EditBox 310, 170, 45, 15, income_rcvd_date
	  EditBox 70, 190, 285, 15, Reason_OP
	  ButtonGroup ButtonPressed
	    OkButton 260, 210, 45, 15
	    CancelButton 310, 210, 45, 15
	  Text 265, 10, 50, 10, "Fraud Referral:"
	  Text 110, 10, 30, 10, "Memb #:"
	  Text 170, 10, 60, 10, "Ot Resp. Memb #:"
	  GroupBox 5, 45, 350, 100, "Overpayment Information"
	  Text 15, 70, 30, 10, "Program:"
	  Text 105, 70, 20, 10, "From:"
	  Text 165, 70, 10, 10, "To:"
	  Text 215, 70, 25, 10, "Claim #"
	  Text 285, 70, 20, 10, "AMT:"
	  Text 15, 90, 30, 10, "Program:"
	  Text 105, 90, 20, 10, "From:"
	  Text 165, 90, 10, 10, "To:"
	  Text 215, 90, 25, 10, "Claim #"
	  Text 285, 90, 20, 10, "AMT:"
	  Text 15, 110, 30, 10, "Program:"
	  Text 105, 110, 20, 10, "From:"
	  Text 165, 110, 10, 10, "To:"
	  Text 215, 110, 25, 10, "Claim #"
	  Text 285, 110, 20, 10, "AMT:"
	  Text 15, 155, 50, 10, "Income Source:"
	  Text 5, 175, 65, 10, "Income Verif Used:"
	  Text 90, 130, 100, 10, "Health Care Resp. Memb(s) #:"
	  Text 230, 130, 75, 10, "Total Federal HC AMT:"
	  Text 25, 195, 40, 10, "OP Reason:"
	  Text 240, 175, 65, 10, "Date Income Rcvd: "
	  Text 130, 55, 30, 10, "(MM/YY)"
	  Text 180, 55, 30, 10, "(MM/YY)"
	  Text 5, 10, 50, 10, "Case Number: "
	  Text 5, 30, 55, 10, "Discovery Date: "
	EndDialog

Do
	err_msg = ""
	dialog overpayment_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
	IF OP_program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
	IF OP_program_II <> "Select:" THEN
		IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF OP_program_III <> "Select:" THEN
		IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verication used for the income recieved. If no verification was received enter N/A."
	'IF isdate(income_rcvd_date) = False or income_rcvd_date = "" then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income recieved."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL check_for_password_without_transmit(are_we_passworded_out)

'----------------------------------------------------------------------------------------------------STAT
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
EMReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
first_name = trim(first_name)
transmit
EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")
'-----------------------------------------------------------------------------------------CASENOTE
'Going to the MISC panel to add claim referral tracking information
Call navigate_to_MAXIS_screen ("STAT", "MISC")
Row = 6
EmReadScreen err_msg, 53, 24, 02
	IF err_msg <> "" THEN
		MsgBox "*** No claim referral can be entered ***" & vbNewLine & err_msg & vbNewLine
	ELSE
        EmReadScreen panel_number, 1, 02, 78
        If panel_number = "0" then
        	EMWriteScreen "NN", 20,79
        	TRANSMIT
        ELSE
        	Do
        		'Checking to see if the MISC panel is empty, if not it will find a new line'
        		EmReadScreen MISC_description, 25, row, 30
        		MISC_description = replace(MISC_description, "_", "")
        		If trim(MISC_description) = "" then
        			PF9
        			EXIT DO
        		Else
        			row = row + 1
        		End if
        	Loop Until row = 17
        	If row = 17 then script_end_procedure("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
        End if
        'writing in the action taken and date to the MISC panel
        EMWriteScreen "Claim Determination", Row, 30
        EMWriteScreen date, Row, 66
        PF3
	    start_a_blank_CASE_NOTE
  	    Call write_variable_in_case_note("-----Claim Referral Tracking-----")
  	    Call write_bullet_and_variable_in_case_note("Program(s)", programs)
  	    Call write_bullet_and_variable_in_case_note("Action Date", date)
  	    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
  	    Call write_variable_in_case_note("-----")
  	    Call write_variable_in_case_note(worker_signature)
	END IF
start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("OVERPAYMENT CLAIM ENTERED" & " (" & first_name & ") " & OP_from & " through " & OP_to)
	CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
	Call write_variable_in_CASE_NOTE("Period " & OP_from & " through " & OP_to)
	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", OP_program)
	CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
	Call write_variable_in_CASE_NOTE("----- ----- -----")
	Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & claim_number & " Amt $" & Claim_amount)
	IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
	IF OP_program_III <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
	CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
	IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
	IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
	IF OP_program = "HC" THEN
		Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
		Call write_bullet_and_variable_in_CASE_NOTE("Health Care claim number", claim_number)
		Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
		Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
	END IF
	CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
	CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
	CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
	CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
	CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	CALL write_variable_in_CASE_NOTE(worker_signature)

	PF3 'to save casenote'

	IF programs = "Health Care" THEN
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
	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & discovery_date & " Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)
	END IF

'---------------------------------------------------------------writing the CCOL case note'
msgbox "Navigating to CCOL to add case note, please contact MiKayla Handley with any concerns."
Call navigate_to_MAXIS_screen("CCOL", "CLSM")
EMWriteScreen claim_number, 4, 9
Transmit
PF4
EMReadScreen existing_case_note, 1, 5, 6
IF existing_case_note = "" THEN
	msgbox "Entering casenote into CCOL"
ELSE
	PF9
END IF
CALL write_variable_in_ccol_note("OVERPAYMENT CLAIM ENTERED" & " (" & first_name & ") " & OP_from & " through " & OP_to)
CALL write_bullet_and_variable_in_ccol_note("Discovery date", discovery_date)
CALL write_variable_in_ccol_note("Period " & OP_from & " through " & OP_to)
CALL write_bullet_and_variable_in_ccol_note("Active Programs", OP_program)
CALL write_bullet_and_variable_in_ccol_note("Source of income", income_source)
CALL write_variable_in_ccol_note("----- ----- -----")
CALL write_variable_in_ccol_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then CALL write_variable_in_ccol_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then CALL write_variable_in_ccol_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF OT_resp_memb <> "" THEN CALL write_bullet_and_variable_in_ccol_note("Other responsible member(s)", OT_resp_memb)
IF EI_checkbox = CHECKED THEN CALL write_variable_in_ccol_note("* Earned Income Disregard Allowed")
IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_ccol_note("* Earned Income Disregard Not Allowed")
IF OP_program = "HC" THEN
	Call write_bullet_and_variable_in_ccol_note("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_ccol_note("HC claim number", claim_number)
	Call write_bullet_and_variable_in_ccol_note("Total federal Health Care amount", Fed_HC_AMT)
	CALL write_variable_in_ccol_note("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
CALL write_bullet_and_variable_in_ccol_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_ccol_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_ccol_note("Date verification received", income_rcvd_date)
CALL write_bullet_and_variable_in_ccol_note("Reason for overpayment", Reason_OP)
CALL write_bullet_and_variable_in_ccol_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_ccol_note("Income verification received", income_rcvd_date)
CALL write_variable_in_ccol_note(" ----- ----- -----")
CALL write_variable_in_ccol_note(worker_signature)
PF3 'exit the case note'
PF3 'back to dail'

script_end_procedure("Overpayment case note entered please review case note to ensure accuracy.")
