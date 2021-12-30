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
'FUNCTIONS LIBRARY BLOCK================================================================================================


'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/18/2020", "GitHub issue #381 Added Requested Claim Adjustment per project request.", "MiKayla Handley")
call changelog_update("03/04/2020", "Removed agency error OP worksheet as the form is now obsolete.", "MiKayla Handley")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
CALL changelog_update("04/15/2019", "Updated script to copy case note to CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Added income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01" 'defaults to 01'
discovery_date = "" & date
servicing_worker = "X127720" ' add the transfer to the end of script '

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog Dialog1, 0, 0, 171, 135, "Overpayment/Claim"
  EditBox 60, 55, 45, 15, MAXIS_case_number
  DropListBox 60, 75, 105, 15, "Select One:"+chr(9)+"Intial Overpayment/Claim"+chr(9)+"Requested Claim Adjustment", claim_actions
  EditBox 60, 95, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 115, 45, 15
    CancelButton 120, 115, 45, 15
    PushButton 110, 55, 55, 15, "CLAIMS", claims_button
  Text 10, 60, 50, 10, "Case Number:"
  Text 10, 80, 50, 10, "Claim Action:"
  Text 10, 100, 40, 10, "Worker Sig:"
  GroupBox 5, 5, 160, 45, "IMPORTANT:"
  Text 10, 15, 150, 10, "CASE/NOTE to be run once claim is complete."
  Text 10, 25, 145, 10, "Does not enter text into CLDL Demand Letter"
  Text 10, 35, 100, 10, "or update worker to X127720"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If ButtonPressed = claims_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/sitepages/Claims_and_Overpayments.aspx")
	Loop until ButtonPressed = -1
	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
    IF claim_actions = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select type of claim action."
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature, for help see utilities. "
	IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

IF claim_actions = "Intial Overpayment/Claim" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 361, 280, "Overpayment Claim Entered"
      EditBox 60, 5, 40, 15, discovery_date
      EditBox 140, 5, 20, 15, memb_number
      EditBox 235, 5, 20, 15, OT_resp_memb
      DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
      DropListBox 50, 40, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
      EditBox 130, 40, 30, 15, OP_from
      EditBox 180, 40, 30, 15, OP_to
      EditBox 245, 40, 35, 15, Claim_number
      EditBox 305, 40, 45, 15, Claim_amount
      DropListBox 50, 60, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
      EditBox 130, 60, 30, 15, OP_from_II
      EditBox 180, 60, 30, 15, OP_to_II
      EditBox 245, 60, 35, 15, Claim_number_II
      EditBox 305, 60, 45, 15, Claim_amount_II
      DropListBox 50, 80, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
      EditBox 130, 80, 30, 15, OP_from_III
      EditBox 180, 80, 30, 15, OP_to_III
      EditBox 245, 80, 35, 15, claim_number_III
      EditBox 305, 80, 45, 15, Claim_amount_III
      DropListBox 50, 100, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_IV
      EditBox 130, 100, 30, 15, OP_from_IV
      EditBox 180, 100, 30, 15, OP_to_IV
      EditBox 245, 100, 35, 15, claim_number_IV
      EditBox 305, 100, 45, 15, Claim_amount_IV
      EditBox 40, 135, 30, 15, HC_from
      EditBox 90, 135, 30, 15, HC_to
      EditBox 160, 135, 50, 15, HC_claim_number
      EditBox 235, 135, 45, 15, HC_claim_amount
      EditBox 40, 155, 30, 15, HC_from_II
      EditBox 90, 155, 30, 15, HC_to_II
      EditBox 160, 155, 50, 15, HC_claim_number_II
      EditBox 235, 155, 45, 15, HC_claim_amount_II
      EditBox 100, 175, 20, 15, HC_resp_memb
      EditBox 235, 175, 45, 15, Fed_HC_AMT
      EditBox 70, 200, 160, 15, income_source
      CheckBox 235, 205, 120, 10, "Earned income disregard allowed", EI_checkbox
      EditBox 70, 220, 160, 15, EVF_used
      EditBox 310, 220, 45, 15, income_rcvd_date
      EditBox 70, 240, 285, 15, Reason_OP
       Text 5, 10, 55, 10, "Discovery date:"
       Text 110, 10, 30, 10, "Memb #:"
       Text 165, 10, 70, 10, "Other resp. memb #:"
       Text 260, 10, 50, 10, "Fraud referral:"
       GroupBox 5, 25, 350, 95, "Overpayment Information"
       Text 130, 30, 30, 10, "(MM/YY)"
       Text 180, 30, 30, 10, "(MM/YY)"
       Text 15, 45, 30, 10, "Program:"
       Text 105, 45, 20, 10, "From:"
       Text 165, 45, 10, 10, "To:"
       Text 215, 45, 25, 10, "Claim #"
       Text 285, 45, 20, 10, "AMT:"
       Text 15, 65, 30, 10, "Program:"
       Text 105, 65, 20, 10, "From:"
       Text 165, 65, 10, 10, "To:"
       Text 215, 65, 25, 10, "Claim #"
       Text 285, 65, 20, 10, "AMT:"
       Text 15, 85, 30, 10, "Program:"
       Text 105, 85, 20, 10, "From:"
       Text 165, 85, 10, 10, "To:"
       Text 215, 85, 25, 10, "Claim #"
       Text 285, 85, 20, 10, "AMT:"
       Text 15, 105, 30, 10, "Program:"
       Text 105, 105, 20, 10, "From:"
       Text 165, 105, 10, 10, "To:"
       Text 215, 105, 25, 10, "Claim #"
       Text 285, 105, 20, 10, "AMT:"
       ButtonGroup ButtonPressed
         OkButton 260, 260, 45, 15
         CancelButton 310, 260, 45, 15
       GroupBox 5, 125, 350, 70, "HC Programs Only"
       Text 15, 160, 20, 10, "From:"
       Text 75, 160, 10, 10, "To:"
       Text 130, 160, 25, 10, "Claim #"
       Text 215, 160, 20, 10, "AMT:"
       Text 15, 140, 20, 10, "From:"
       Text 75, 140, 10, 10, "To:"
       Text 130, 140, 25, 10, "Claim #"
       Text 215, 140, 20, 10, "AMT:"
       Text 15, 180, 80, 10, "HC OT resp. Memb(s) #:"
       Text 160, 180, 75, 10, "Total federal HC AMT:"
       Text 5, 205, 50, 10, "Income source:"
       Text 5, 225, 65, 10, "Income verif used:"
       Text 5, 245, 60, 10, "Reason for claim:"
       Text 235, 225, 75, 10, "Date income received:"
     EndDialog

    Do
        Do
        	err_msg = ""
        	dialog Dialog1
        	cancel_confirmation
			IF memb_number = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the member number."
        	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
        	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
        	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
        	IF OP_program_II <> "Select:" THEN
				IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	   		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
        		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
        	END IF
    	    IF OP_program_III <> "Select:" THEN
				IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	    	IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
    	    	IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
    	    END IF
    	    IF OP_program_IV <> "Select:" THEN
				IF OP_from_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	    	IF Claim_number_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
    	    	IF Claim_amount_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
    	    END IF
            	IF HC_claim_number <> "" THEN
            		IF HC_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment started."
            		IF HC_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment ended."
            		IF HC_claim_amount = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
            	END IF
            IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verification used for the income received. If no verification was received enter N/A."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = False
	'---------------------------------------------------------------------------------------------'client information
	back_to_self
	CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
	IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
    EMwritescreen MEMB_number, 20, 76
	TRANSMIT
	EMReadscreen panel_MEMB_number, 2, 4, 33
	'MsgBox panel_MEMB_number & " ~ " &  MEMB_number
	IF panel_MEMB_number <> MEMB_number THEN script_end_procedure_with_error_report("This MEMB was not found, the script will now end.")
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
	last_name = trim(replace(last_name, "_", ""))
	first_name = trim(replace(first_name, "_", ""))
	mid_initial = replace(mid_initial, "_", "")
	client_name = MEMB_number & " - " & last_name &  ", " & first_name & " " & mid_initial
    client_name = trim(client_name)
    '-----------------------------------------------------------------------------------------CASENOTE
    IF OP_program = "FS" or OP_program_II = "FS" or OP_program_III = "FS" or OP_program_IV = "FS" or OP_program = "MF" or OP_program_II = "MF" or OP_program_III = "MF" or OP_program_IV = "MF" THEN
    	'Going to the MISC panel to add claim referral tracking information
    	Call navigate_to_MAXIS_screen ("STAT", "MISC")
    	Row = 6
    	EmReadScreen panel_number, 1, 02, 73
    	If panel_number = "0" then
    		EMWriteScreen "NN", 20,79
    		TRANSMIT
    		EmReadScreen MISC_error_check,  74, 24, 02
    		IF trim(MISC_error_check) = "" THEN
    			case_note_only = FALSE
    		else
    			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion, "Message handling")
    			IF maxis_error_check = vbYes THEN
    				case_note_only = TRUE 'this will case note only'
    			END IF
    			IF maxis_error_check= vbNo THEN
    				case_note_only = FALSE 'this will update the panels and case note'
    			END IF
    		END IF
    	END IF

    	Do
    		'Checking to see if the MISC panel is empty, if not it will find a new line'
    		EmReadScreen MISC_description, 25, row, 30
    		MISC_description = replace(MISC_description, "_", "")
    		If trim(MISC_description) = "" then
    			'PF9
    			EXIT DO
    		Else
    			row = row + 1
    		End if
    	Loop Until row = 17
    	IF row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")

    	'writing in the action taken and date to the MISC panel
    	PF9
    	EMWriteScreen "Determination-OP Entered", Row, 30
    	EMWriteScreen date, Row, 66
    	TRANSMIT

    	start_a_blank_case_note
    	Call write_variable_in_case_note("-----Claim Referral Tracking - Claim Determination-----")
    	IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
    	Call write_bullet_and_variable_in_case_note("Action Date", date)
    	Call write_bullet_and_variable_in_case_note("Program(s)", programs)
    	Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    	Call write_variable_in_case_note(worker_signature)
    	PF3
    END IF

    '-----------------------------------------------------------------------------------------CASENOTE
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("OVERPAYMENT CLAIM ENTERED" & " (" & client_name & ") ")
    CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
    Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
    IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
    IF OP_program_III <> "Select:" then	Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
    IF HC_claim_number <> "" THEN
    	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amount $" & HC_Claim_amount)
    	Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
    	Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
    	Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF HC_claim_number_II <> "" THEN
    	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " Amount $" & HC_claim_amount_II)
    	Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
    	Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
    	Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
    IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
    CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
    CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
    CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
    CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
    CALL write_variable_in_CASE_NOTE("----- ----- -----")
    CALL write_variable_in_CASE_NOTE(worker_signature)
    PF3 'to save casenote'

    IF HC_claim_number <> "" THEN
    	EMWriteScreen "x", 5, 3
    	TRANSMIT
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
    	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & discovery_date & "HC Claim # " & HC_claim_number, "CASE NOTE" & vbcr & message_array,"", False)
    END IF

    '---------------------------------------------------------------writing the CCOL case note'
    Call navigate_to_MAXIS_screen("CCOL", "CLSM")
    EMWriteScreen Claim_number, 4, 9
    TRANSMIT
    EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
    error_check = trim(error_check)
    If error_check <> "" then script_end_procedure_with_error_report(error_check & ". Unable to update this case. Please review case, and run the script again if applicable.")
    PF4
    EMReadScreen existing_case_note, 1, 5, 6
    IF existing_case_note = "" THEN
    	PF4
    ELSE
    	PF9
    END IF

    Call write_variable_in_CCOL_note("OVERPAYMENT CLAIM ENTERED" & " (" & client_name & ") ")
    CALL write_bullet_and_variable_in_CCOL_note("Discovery date", discovery_date)
    CALL write_bullet_and_variable_in_CCOL_note("Active Programs", programs)
    CALL write_bullet_and_variable_in_CCOL_note("Source of income", income_source)
    IF OP_program <> "Select:" then Call write_variable_in_CCOL_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
    IF OP_program_II <> "Select:" then Call write_variable_in_CCOL_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
    IF OP_program_III <> "Select:" then Call write_variable_in_CCOL_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" then Call write_variable_in_CCOL_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
    IF HC_claim_number <> "" THEN
    	Call write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
    	Call write_bullet_and_variable_in_CCOL_note("Health Care responsible members", HC_resp_memb)
    	Call write_bullet_and_variable_in_CCOL_note("Total Federal Health Care amount", Fed_HC_AMT)
    	Call write_variable_in_CCOL_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF HC_claim_number_II <> "" THEN
    	Call write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " Amount $" & HC_claim_amount_II)
    	Call write_bullet_and_variable_in_CCOL_note("Health Care responsible members", HC_resp_memb)
    	Call write_bullet_and_variable_in_CCOL_note("Total Federal Health Care amount", Fed_HC_AMT)
    	Call write_variable_in_CCOL_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Allowed")
    IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Not Allowed")
    CALL write_bullet_and_variable_in_CCOL_note("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_CCOL_note("Income verification received", EVF_used)
    CALL write_bullet_and_variable_in_CCOL_note("Date verification received", income_rcvd_date)
    CALL write_bullet_and_variable_in_CCOL_note("Reason for overpayment", Reason_OP)
    CALL write_bullet_and_variable_in_CCOL_note("Other responsible member(s)", OT_resp_memb)
    CALL write_variable_in_CCOL_note("----- ----- -----")
    CALL write_variable_in_CCOL_note(worker_signature)
    PF3 'to save casenote'
END IF

IF claim_actions = "Requested Claim Adjustment" THEN
    BeginDialog Dialog1, 0, 0, 221, 165, "Requested Claim Adjustment"
      EditBox 65, 5, 50, 15, claim_number
      EditBox 75, 25, 40, 15, original_claim_amount
      EditBox 175, 25, 40, 15, corrected_claim_amount
      EditBox 75, 45, 40, 15, adjustment_amount
      EditBox 125, 65, 35, 15, OP_from
      EditBox 180, 65, 35, 15, OP_to
      EditBox 90, 85, 125, 15, reason_correction
      EditBox 90, 105, 125, 15, requested_verif
      EditBox 50, 125, 165, 15, other_notes
      CheckBox 170, 10, 50, 10, "MFIP Claim", MFIP_Claim_checkbox
      ButtonGroup ButtonPressed
        OkButton 115, 145, 50, 15
        CancelButton 170, 145, 45, 15
      Text 5, 130, 45, 10, "Other notes:"
      Text 100, 70, 20, 10, "From:"
      Text 165, 70, 15, 10, "To:"
      Text 5, 90, 75, 10, "Reason for correction:"
      Text 130, 55, 35, 10, "(MM/YY)"
      Text 5, 10, 50, 10, "Claim number:"
      Text 185, 55, 35, 10, "(MM/YY)"
      Text 120, 30, 55, 10, "Correct Amount:"
      Text 5, 50, 65, 10, "Adjustment Amount:"
      Text 5, 70, 50, 10, "Correct Period"
      Text 5, 30, 55, 10, "Original Amount:"
      Text 5, 110, 80, 10, "Requested verifications:"
    EndDialog

	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
	      	IF IsNumeric(claim_number) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid claim number."
			IF IsNumeric(original_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid original claim amount(do not include $)."
	        IF IsNumeric(corrected_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid original claim amount(do not include $)."
			IF OP_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
			IF OP_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
			IF reason_correction = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the reason the correction is needed. "
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false

	IF MFIP_Claim_checkbox = CHECKED THEN
	    BeginDialog Dialog1, 0, 0, 276, 125, "MFIP Amount Adjustment "
	      EditBox 90, 20, 40, 15, tanf_elig_cash
	      EditBox 90, 40, 40, 15, tanf_housing_grant
	      EditBox 90, 60, 40, 15, federal_food
	      EditBox 230, 20, 40, 15, state_funds_cash
	      EditBox 230, 40, 40, 15, state_housing_grant
	      EditBox 230, 60, 40, 15, state_food
	      EditBox 135, 85, 40, 15, total_amount
	      ButtonGroup ButtonPressed
	        OkButton 170, 105, 50, 15
	        CancelButton 225, 105, 45, 15
	      Text 140, 65, 60, 10, "STATE FOOD:"
	      Text 105, 90, 30, 10, "TOTAL:"
	      Text 85, 5, 90, 10, "MFIP AMOUNT"
	      Text 140, 25, 75, 10, "STATE FUNDS CASH:"
	      Text 5, 45, 85, 10, "TANF HOUSING GRANT:"
	      Text 5, 65, 60, 10, "FEDERAL FOOD:"
	      Text 5, 25, 65, 10, "TANF ELIG CASH:"
	      Text 140, 45, 90, 10, "STATE HOUSING GRANT:"
	    EndDialog

		Do
			Do
				err_msg = ""
				Dialog Dialog1
				cancel_without_confirmation
		      	IF IsNumeric(total_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the total amount."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false
	END IF

	'-----------------------------------------------------------------------------------------CASENOTE
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("Requested Claim Adjustment")
    CALL write_variable_in_CASE_NOTE("* Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number)
	CALL write_bullet_and_variable_in_CASE_NOTE("Original Amount", original_claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Correct Amount",  corrected_claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Adjustment Amount",  adjustment_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verif)
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("TANF Cash", tanf_elig_cash)
	Call write_bullet_and_variable_in_CASE_NOTE("TANF Housing Grant", tanf_housing_grant)
	Call write_bullet_and_variable_in_CASE_NOTE("Federal Food", federal_food)
	Call write_bullet_and_variable_in_CASE_NOTE("State Cash", state_funds_cash)
	Call write_bullet_and_variable_in_CASE_NOTE("State Housing", state_housing_grant)
	Call write_bullet_and_variable_in_CASE_NOTE("State Food", state_food)
	Call write_bullet_and_variable_in_CASE_NOTE("Total", total_amount)
	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3

	EMWriteScreen "x", 5, 3
	TRANSMIT
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

	'---------------------------------------------------------------writing the CCOL case note'
    Call navigate_to_MAXIS_screen("CCOL", "CLSM")
    EMWriteScreen Claim_number, 4, 9
    TRANSMIT
    EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
    error_check = trim(error_check)
    If error_check <> "" then script_end_procedure_with_error_report(error_check & ". Unable to update CCOL. Please review case, and run the script again if applicable.")
    PF4
    EMReadScreen existing_case_note, 1, 5, 6
    IF existing_case_note = "" THEN
    	PF4
    ELSE
    	PF9
    END IF

	Call write_variable_in_CCOL_note("Requested Claim Adjustment")
	CALL write_variable_in_CCOL_note("* Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number)
	CALL write_bullet_and_variable_in_CCOL_note("Original Amount", original_claim_amount)
	Call write_bullet_and_variable_in_CCOL_note("Correct Amount", corrected_claim_amount)
	Call write_bullet_and_variable_in_CCOL_note("Adjustment Amount",  adjustment_amount)
	Call write_bullet_and_variable_in_CCOL_note("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CCOL_note("Requested verifications", requested_verif)
	Call write_bullet_and_variable_in_CCOL_note("Other notes", other_notes)
	Call write_bullet_and_variable_in_CCOL_note("TANF Cash", tanf_elig_cash)
	Call write_bullet_and_variable_in_CCOL_note("TANF Housing Grant", tanf_housing_grant)
	Call write_bullet_and_variable_in_CCOL_note("Federal Food", federal_food)
	Call write_bullet_and_variable_in_CCOL_note("State Cash", state_funds_cash)
	Call write_bullet_and_variable_in_CCOL_note("State Housing", state_housing_grant)
	Call write_bullet_and_variable_in_CCOL_note("State Food", state_food)
	Call write_bullet_and_variable_in_CCOL_note("Total", total_amount)
	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	CALL write_variable_in_CCOL_note(worker_signature)
	PF3
	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "","Requested Claim Adjustment " &  MAXIS_case_number & " Member # " & memb_number & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number, "CASE NOTE" & vbcr & message_array,"", False)
END IF
script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL please review case note to ensure accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/29/2021
'--Tab orders reviewed & confirmed----------------------------------------------12/29/2021
'--Mandatory fields all present & Reviewed--------------------------------------12/29/2021
'--All variables in dialog match mandatory fields-------------------------------12/29/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/29/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------12/29/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------12/29/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/29/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------12/29/2021
'--PRIV Case handling reviewed -------------------------------------------------12/29/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/29/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/29/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------12/29/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------12/29/2021
'--comment Code-----------------------------------------------------------------12/29/2021
'--Update Changelog for release/update------------------------------------------12/29/2021
'--Remove testing message boxes-------------------------------------------------12/29/2021
'--Remove testing code/unnecessary code-----------------------------------------12/29/2021
'--Review/update SharePoint instructions----------------------------------------12/29/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------12/29/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------12/29/2021
