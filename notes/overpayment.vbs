'GATHERING STATS===========================================================================================
name_of_script = "NOTES - OVERPAYMENT.vbs"
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
call changelog_update("01/29/2023", "Enhanced claim information to ensure that decimal points are added to claim amounts, and not added to claim #'s.", "Ilse Ferris, Hennepin County")
call changelog_update("11/17/2022", "Updated bug in script where claim # and claim amt were transposed.", "Ilse Ferris, Hennepin County")
call changelog_update("11/15/2022", "Enhanced CCOL Notes to make notes in all claims vs. only 1st claim. Updated background functioning.", "Ilse Ferris, Hennepin County")
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
closing_message = "Overpayment case note entered, copied to CCOL and the claims team has been emailed. Please review case & claim notes to ensure accuracy."

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 171, 135, "Overpayment/Claims"
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
        Do
            err_msg = ""
	    	Dialog Dialog1
	    	cancel_without_confirmation
	    	If ButtonPressed = claims_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Claims.aspx")
	    Loop until ButtonPressed = -1
	    Call validate_MAXIS_case_number(err_msg, "*")
        IF claim_actions = "Select One:" then err_msg = err_msg & vbNewLine & "* Select type of claim action."
	    IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
	    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

IF claim_actions = "Intial Overpayment/Claim" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 361, 320, "Overpayment Claim Enty"
        EditBox 60, 10, 40, 15, discovery_date
        EditBox 140, 10, 20, 15, memb_number
        EditBox 235, 10, 20, 15, OT_resp_memb
        DropListBox 310, 10, 45, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", fraud_referral
        DropListBox 60, 60, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
        EditBox 130, 60, 30, 15, OP_from
        EditBox 180, 60, 30, 15, OP_to
        EditBox 235, 60, 35, 15, Claim_amount
        EditBox 285, 60, 45, 15, Claim_number
        DropListBox 60, 80, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
        EditBox 130, 80, 30, 15, OP_from_II
        EditBox 180, 80, 30, 15, OP_to_II
        EditBox 235, 80, 35, 15, Claim_amount_II
        EditBox 285, 80, 45, 15, Claim_number_II
        DropListBox 60, 100, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
        EditBox 130, 100, 30, 15, OP_from_III
        EditBox 180, 100, 30, 15, OP_to_III
        EditBox 235, 100, 35, 15, Claim_amount_III
        EditBox 285, 100, 45, 15, claim_number_III
        DropListBox 60, 120, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_IV
        EditBox 130, 120, 30, 15, OP_from_IV
        EditBox 180, 120, 30, 15, OP_to_IV
        EditBox 235, 120, 35, 15, Claim_amount_IV
        EditBox 285, 120, 45, 15, claim_number_IV
        EditBox 130, 165, 30, 15, HC_from
        EditBox 180, 165, 30, 15, HC_to
        EditBox 235, 165, 35, 15, HC_claim_amount
        EditBox 285, 165, 45, 15, HC_claim_number
        EditBox 130, 185, 30, 15, HC_from_II
        EditBox 180, 185, 30, 15, HC_to_II
        EditBox 235, 185, 35, 15, HC_claim_amount_II
        EditBox 285, 185, 45, 15, HC_claim_number_II
        EditBox 130, 205, 20, 15, HC_resp_memb
        EditBox 235, 205, 35, 15, Fed_HC_AMT
        EditBox 70, 240, 160, 15, income_source
        DropListBox 310, 240, 45, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", EI_disregard
        EditBox 70, 260, 160, 15, EVF_used
        EditBox 310, 260, 45, 15, income_rcvd_date
        EditBox 70, 280, 285, 15, Reason_OP
        ButtonGroup ButtonPressed
            OkButton 260, 300, 45, 15
            CancelButton 310, 300, 45, 15
        Text 110, 15, 30, 10, "Memb #:"
        GroupBox 5, 30, 350, 205, "Overpayment Information"
        Text 260, 15, 50, 10, "Fraud referral:"
        Text 165, 15, 70, 10, "Other resp. memb #:"
        Text 40, 210, 85, 10, "HC other resp. Memb(s)#:"
        Text 160, 210, 75, 10, "Total federal HC AMT:"
        Text 20, 245, 50, 10, "Income source:"
        Text 5, 15, 55, 10, "Discovery date:"
        Text 235, 245, 75, 10, "EI Disregard Allowed?:"
        Text 10, 265, 60, 10, "Income verif used:"
        Text 10, 285, 60, 10, "Reason for claim:"
        Text 235, 265, 75, 10, "Date income received:"
        Text 35, 170, 90, 10, "Health Care Only - Claim 1:"
        Text 35, 190, 90, 10, "Health Care Only - Claim 2:"
        Text 25, 65, 35, 10, "1st Claim:"
        Text 15, 45, 320, 10, "       Claim          Program              Start (MM/YY) - End (MM/YY)          Amount $              Claim # "
        Text 25, 85, 35, 10, "2nd Claim:"
        Text 25, 125, 35, 10, "4th Claim:"
        Text 25, 105, 35, 10, "3rd Claim:"
        GroupBox 10, 150, 325, 75, "Health Care Only"
    EndDialog

    Do
        Do
        	err_msg = ""
        	dialog Dialog1
        	cancel_confirmation
			IF memb_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the member number."
        	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* Select a fraud referral entry."
            If OP_program = "Select:" then
                If trim(HC_claim_number) = "" THEN err_msg = err_msg & vbNewLine & "* You must enter at least claim or health care claim."
            End if
            '1st Claim
            IF OP_program <> "Select:" THEN
                IF trim(OP_from) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 1st overpayment period period (MM/YY)."
                IF trim(OP_to) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 1st overpayment period period (MM/YY)."
                IF trim(Claim_amount) = "" or instr(Claim_amount, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 1st claim amount (including decimal point.)"
                IF trim(Claim_number) = "" or instr(claim_number, ".") <> 0 THEN err_msg = err_msg & vbNewLine & "* Enter the 1st claim number."
            END IF
            '2nd Claim
        	IF OP_program_II <> "Select:" THEN
				IF trim(OP_from_II) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 2nd overpayment period period (MM/YY)."
				IF trim(OP_to_II) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 2nd overpayment period period (MM/YY)."
                IF trim(Claim_amount_II) = "" or instr(Claim_amount_II, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 2nd claim amount (including decimal point.)"
                IF trim(Claim_number_II) = "" or instr(claim_number_II, ".") <> 0 THEN err_msg = err_msg & vbNewLine & "* Enter the 2nd claim number."
        	END IF
    	    IF OP_program_III <> "Select:" THEN
				IF trim(OP_from_III) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 3rd overpayment period period (MM/YY)."
				IF trim(OP_to_III) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 3rd overpayment period period (MM/YY)."
    	    	IF trim(Claim_amount_III) = "" or instr(Claim_amount_III, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 3rd claim amount (including decimal point.)"
                IF trim(Claim_number_III) = "" or instr(claim_number_III, ".") <> 0 THEN err_msg = err_msg & vbNewLine & "* Enter the 3rd claim number."
    	    END IF
    	    IF OP_program_IV <> "Select:" THEN
				IF trim(OP_from_IV) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 4th overpayment period period (MM/YY)."
				IF trim(OP_to_IV) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 4th overpayment period period (MM/YY)."
    	    	IF trim(Claim_amount_IV) = "" or instr(Claim_amount_IV, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 4th claim amount (including decimal point.)"
                IF trim(Claim_number_IV) = "" or instr(claim_number_IV, ".") <> 0 THEN err_msg = err_msg & vbNewLine & "* Enter the 4th claim number."
    	    END IF
            IF trim(HC_claim_number) <> "" THEN
                If instr(HC_claim_number, ".") <> 0 then err_msg = err_msg & vbNewLine & "* Review 1st Health Care claim #, remove the decimal point if applicable."
            	IF trim(HC_from) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 1st Health Care overpayment period period (MM/YY)."
            	IF trim(HC_to) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 1st Health Care overpayment period period (MM/YY)."
            	IF trim(HC_claim_amount) = "" or instr(HC_claim_amount, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 1st Health Care claim amount (including decimal point.)"
            END IF
            IF trim(HC_claim_number_II) <> "" THEN
                If instr(HC_claim_number_II, ".") <> 0 then err_msg = err_msg & vbNewLine & "* Review 2nd Health Care claim #, remove the decimal point if applicable."
                IF trim(HC_from_II) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the start of the 2nd Health Care overpayment period period (MM/YY)."
                IF trim(HC_to_II) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the end of the 2nd Health Care overpayment period period (MM/YY)."
                IF trim(HC_claim_amount_II) = "" or instr(HC_claim_amount_II, ".") = 0 then err_msg = err_msg & vbNewLine & "* Enter the 2nd Health Care claim amount (including decimal point.)"
            END IF
            If EI_disregard = "Select:" THEN err_msg = err_msg & vbNewLine & "* Was an earned income disregard allowed in the overpayment?"
            IF trim(EVF_used) = "" then err_msg = err_msg & vbNewLine & "* Enter verification used for the income received. If no verification was received enter N/A."
            IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* Enter a reason for the overpayment in as much detail as possible."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = False

    'Creating an array of all the claims to claim note for each one
    all_claim_numbers = ""
    If trim(Claim_number) <> "" then all_claim_numbers = all_claim_numbers & trim(Claim_number) & "|"
    If trim(Claim_number_II) <> "" then all_claim_numbers = all_claim_numbers & trim(Claim_number_II) & "|"
    If trim(Claim_number_III) <> "" then all_claim_numbers = all_claim_numbers & trim(Claim_number_III) & "|"
    If trim(Claim_number_IV) <> "" then all_claim_numbers = all_claim_numbers & trim(Claim_number_IV) & "|"
    If trim(HC_claim_number) <> "" then all_claim_numbers = all_claim_numbers & trim(HC_claim_number) & "|"
    If trim(HC_claim_number_II) <> "" then all_claim_numbers = all_claim_numbers & trim(HC_claim_number_II) & "|"

    claim_array = split(all_claim_numbers, "|")

	'---------------------------------------------------------------------------------------------'client information
	back_to_self
	CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
	IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
    Call write_value_and_transmit(MEMB_number, 20, 76)

	EMReadscreen panel_MEMB_number, 2, 4, 33
	IF panel_MEMB_number <> MEMB_number THEN script_end_procedure_with_error_report("This MEMB was not found, the script will now end.")

	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
	last_name = trim(replace(last_name, "_", ""))
	first_name = trim(replace(first_name, "_", ""))
	mid_initial = replace(mid_initial, "_", "")
	client_name = MEMB_number & " - " & last_name &  ", " & first_name & " " & mid_initial
    client_name = trim(client_name)

    Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

    '----------------------------------------------------------------------------------------Claim Referral Tracking for MFIP or SNAP
    IF OP_program = "FS" or OP_program_II = "FS" or OP_program_III = "FS" or OP_program_IV = "FS" or OP_program = "MF" or OP_program_II = "MF" or OP_program_III = "MF" or OP_program_IV = "MF" THEN
    	'Going to the MISC panel to add claim referral tracking information
    	Call navigate_to_MAXIS_screen ("STAT", "MISC")
    	Row = 6
    	EmReadScreen panel_number, 1, 02, 73
    	If panel_number = "0" then
    		Call write_value_and_transmit("NN", 20, 79)
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
    			EXIT DO
    		Else
    			row = row + 1
    		End if
    	Loop Until row = 17
    	IF row = 17 then
            closing_message = closing_message & vbcr & vbcr & "The script was unable to update the MISC panel. Delete a line(s) on STAT/MISC, and run script again or update manually."
            case_note_only = True 'can only case note if all the MISC panels are all filled.
        Else
    	    'writing in the action taken and date to the MISC panel
    	    PF9
    	    EMWriteScreen "Determination-OP Entered", Row, 30
    	    Call write_value_and_transmit(date, Row, 66)
        End if

    	start_a_blank_case_note
    	Call write_variable_in_case_note("-----Claim Referral Tracking - Claim Determination-----")
    	IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
    	Call write_bullet_and_variable_in_case_note("Action Date", date)
    	Call write_bullet_and_variable_in_case_note("Active Program(s)", list_active_programs)
    	Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
        Call write_variable_in_case_note("---")
    	Call write_variable_in_case_note(worker_signature)
    	PF3
    END IF
    '-----------------------------------------------------------------------------------------CASE/NOTE
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("OVERPAYMENT CLAIM ENTERED" & " (" & client_name & ") ")
    CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Active Program(s)", list_active_programs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
    Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " for $" & Claim_amount)
    IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " for $" & Claim_amount_II)
    IF OP_program_III <> "Select:" then	Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " for $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " for $" & Claim_amount_IV)
    'health care
    IF HC_claim_number <> "" THEN Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " for $" & HC_Claim_amount)
    IF HC_claim_number_II <> "" THEN Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " for $" & HC_claim_amount_II)
    Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
    Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
    If HC_claim_number <> "" or HC_claim_number_II <> "" then
        Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
        CALL write_variable_in_CASE_NOTE("---")
    End if
    'Income and OP reasons/info
    Call write_bullet_and_variable_in_CASE_NOTE("Earned Income Disregard Applied?", EI_disregard)
    CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
    CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
    CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
    CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)
    PF3 'to save casenote'

    IF HC_claim_number <> "" THEN
    	Call write_value_and_transmit("X", 5, 3)
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
    For each unique_claim in claim_array
        If trim(unique_claim) = "" then exit for
        Call navigate_to_MAXIS_screen("CCOL", "CLSM")
        Call write_value_and_transmit(unique_claim, 4, 9)

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
        CALL write_bullet_and_variable_in_CCOL_note("Active Program(s)", list_active_programs)
        CALL write_bullet_and_variable_in_CCOL_note("Source of income", income_source)
        Call write_variable_in_CCOL_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " for $" & Claim_amount)
        IF OP_program_II <> "Select:" then Call write_variable_in_CCOL_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " for $" & Claim_amount_II)
        IF OP_program_III <> "Select:" then	Call write_variable_in_CCOL_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " for $" & Claim_amount_III)
        IF OP_program_IV <> "Select:" then Call write_variable_in_CCOL_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " for $" & Claim_amount_IV)
        'health care
        IF HC_claim_number <> "" THEN Call write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " for $" & HC_Claim_amount)
        IF HC_claim_number_II <> "" THEN Call write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " for $" & HC_claim_amount_II)
        Call write_bullet_and_variable_in_CCOL_note("Health Care responsible members", HC_resp_memb)
        Call write_bullet_and_variable_in_CCOL_note("Total Federal Health Care amount", Fed_HC_AMT)
        If HC_claim_number <> "" or HC_claim_number_II <> "" then
            Call write_variable_in_CCOL_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
            CALL write_variable_in_CCOL_note("---")
        End if
        'Income and OP reasons/info
        Call write_bullet_and_variable_in_CCOL_note("Earned Income Disregard Applied?", EI_disregard)
        CALL write_bullet_and_variable_in_CCOL_note("Fraud referral made", fraud_referral)
        CALL write_bullet_and_variable_in_CCOL_note("Income verification received", EVF_used)
        CALL write_bullet_and_variable_in_CCOL_note("Date verification received", income_rcvd_date)
        CALL write_bullet_and_variable_in_CCOL_note("Reason for overpayment", Reason_OP)
        CALL write_bullet_and_variable_in_CCOL_note("Other responsible member(s)", OT_resp_memb)
        CALL write_variable_in_CCOL_note("---")
        CALL write_variable_in_CCOL_note(worker_signature)
        PF3 'to save claim notes
    Next
END IF

IF claim_actions = "Requested Claim Adjustment" THEN
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 221, 165, "Requested Claim Adjustment"
      EditBox 65, 5, 50, 15, claim_number
      EditBox 75, 25, 40, 15, original_claim_amount
      EditBox 175, 25, 40, 15, correct_claim_amount
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
	      	IF IsNumeric(claim_number) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid claim number."
			IF IsNumeric(original_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid original claim amount(do not include $)."
	        IF IsNumeric(correct_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid original claim amount(do not include $)."
            If trim(adjustment_amount) <> "" then
                If IsNumeric(adjustment_amount) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid adjusted claim amount(do not include $)."
            End if
			IF trim(OP_from) = "" THEN err_msg = err_msg & vbNewLine &  "* Enter the start date of the correction, month and year (MM/YY)."
			IF trim(OP_to) = "" THEN err_msg = err_msg & vbNewLine &  "* Enter the end date of the correction, month and year (MM/YY)."
			IF trim(reason_correction) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the reason the correction is needed."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false

	IF MFIP_Claim_checkbox = CHECKED THEN
        Dialog1 = ""
	    BeginDialog Dialog1, 0, 0, 276, 125, "MFIP Amount Adjustment"
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
                If trim(tanf_elig_cash) <> "" then
                    if IsNumeric(tanf_elig_cash) = False then err_msg = err_msg & vbcr & "* Enter a numeric TANF Cash amount."
                End if
                If trim(tanf_housing_grant) <> "" then
                    if IsNumeric(tanf_housing_grant) = False then err_msg = err_msg & vbcr & "* Enter a numeric TANF Housing Grant amount."
                End if
                If trim(federal_food) <> "" then
                    if IsNumeric(federal_food) = False then err_msg = err_msg & vbcr & "* Enter a numeric Federal Food amount."
                End if
                If trim(state_funds_cash) <> "" then
                    if IsNumeric(state_funds_cash) = False then err_msg = err_msg & vbcr & "* Enter a numeric State Cash amount."
                End if
                If trim(state_housing_grant) <> "" then
                    if IsNumeric(state_housing_grant) = False then err_msg = err_msg & vbcr & "* Enter a numeric State Housing Grant amount."
                End if
                If trim(state_food) <> "" then
                    if IsNumeric(state_food) = False then err_msg = err_msg & vbcr & "* Enter a numeric State Food amount."
                End if
                If trim(total_amount) <> "" then
                    if IsNumeric(total_amount) = False then err_msg = err_msg & vbcr & "* Enter a numeric total amount."
		        End if
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
	Call write_bullet_and_variable_in_CASE_NOTE("Correct Amount", correct_claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Adjustment Amount", adjustment_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CASE_NOTE("Requested verifications", requested_verif)
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    IF MFIP_Claim_checkbox = checked then
        CALL write_variable_in_CASE_NOTE("---")
	    Call write_bullet_and_variable_in_CASE_NOTE("TANF Cash", tanf_elig_cash)
	    Call write_bullet_and_variable_in_CASE_NOTE("TANF Housing Grant", tanf_housing_grant)
	    Call write_bullet_and_variable_in_CASE_NOTE("Federal Food", federal_food)
	    Call write_bullet_and_variable_in_CASE_NOTE("State Cash", state_funds_cash)
	    Call write_bullet_and_variable_in_CASE_NOTE("State Housing", state_housing_grant)
	    Call write_bullet_and_variable_in_CASE_NOTE("State Food", state_food)
	    Call write_bullet_and_variable_in_CASE_NOTE("Total", total_amount)
    End if
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3

	Call write_value_and_transmit("X", 5, 3)
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
				PF8												'goes to the next line and resets the row to read'
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages

	'---------------------------------------------------------------writing the CCOL case note'
    Call navigate_to_MAXIS_screen("CCOL", "CLSM")
    Call write_value_and_transmit(Claim_number, 4, 9)

    EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
    error_check = trim(error_check)
    If error_check <> "" then script_end_procedure_with_error_report(error_check & ". Unable to CCOL. Please review claim case, and run the script again if applicable.")
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
	Call write_bullet_and_variable_in_CCOL_note("Correct Amount", correct_claim_amount)
	Call write_bullet_and_variable_in_CCOL_note("Adjustment Amount",  adjustment_amount)
	Call write_bullet_and_variable_in_CCOL_note("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CCOL_note("Requested verifications", requested_verif)
	Call write_bullet_and_variable_in_CCOL_note("Other notes", other_notes)
    IF MFIP_Claim_checkbox = checked then
        CALL write_variable_in_CCOL_note("---")
	    Call write_bullet_and_variable_in_CCOL_note("TANF Cash", tanf_elig_cash)
	    Call write_bullet_and_variable_in_CCOL_note("TANF Housing Grant", tanf_housing_grant)
	    Call write_bullet_and_variable_in_CCOL_note("Federal Food", federal_food)
	    Call write_bullet_and_variable_in_CCOL_note("State Cash", state_funds_cash)
	    Call write_bullet_and_variable_in_CCOL_note("State Housing", state_housing_grant)
	    Call write_bullet_and_variable_in_CCOL_note("State Food", state_food)
	    Call write_bullet_and_variable_in_CCOL_note("Total", total_amount)
    End if
	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	CALL write_variable_in_CCOL_note(worker_signature)
	PF3
	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "","Requested Claim Adjustment " &  MAXIS_case_number & " Member # " & memb_number & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number, "CASE NOTE" & vbcr & message_array,"", False)
END IF

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/15/2022
'--Tab orders reviewed & confirmed----------------------------------------------11/15/2022
'--Mandatory fields all present & Reviewed--------------------------------------11/15/2022
'--All variables in dialog match mandatory fields-------------------------------11/15/2022
'Review dialog names for content and content fit in dialog----------------------01/29/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------11/15/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------11/15/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------11/15/2022--------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/15/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------11/15/2022--------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------11/15/2022--------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------11/15/2022
'--Out-of-County handling reviewed----------------------------------------------11/15/2022--------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/15/2022
'--BULK - review output of statistics and run time/count (if applicable)--------11/15/2022--------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------11/15/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/15/2022
'--Incrementors reviewed (if necessary)-----------------------------------------11/15/2022
'--Denomination reviewed -------------------------------------------------------11/15/2022
'--Script name reviewed---------------------------------------------------------11/15/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------11/15/2022--------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/15/2022
'--comment Code-----------------------------------------------------------------11/15/2022
'--Update Changelog for release/update------------------------------------------11/15/2022
'--Remove testing message boxes-------------------------------------------------11/15/2022
'--Remove testing code/unnecessary code-----------------------------------------11/15/2022
'--Review/update SharePoint instructions----------------------------------------11/15/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------11/15/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/15/2022
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------01/29/2023
'--Complete misc. documentation (if applicable)---------------------------------11/15/2022
'--Update project team/issue contact (if applicable)----------------------------11/15/2022
