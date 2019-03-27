name_of_script = "ACTIONS - DEU-PARIS MATCH CLEARED CC.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 700         'manual run time in seconds
STATS_denomination = "C"      'C is for each case
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()
CALL changelog_update("09/28/2018", "Added handling for more than two states of PARIS matches on INSM.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updates made to correct error.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My DOcuments folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
BeginDialog case_number_dialog, 0, 0, 131, 65, "Dialog"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 60, 25, 30, 15, MEMB_number
  ButtonGroup ButtonPressed
    OkButton 20, 45, 50, 15
    CancelButton 75, 45, 50, 15
  Text 5, 30, 55, 10, "MEMB Number:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

'---------------------------------------------------------------------THE SCRIPT
'Connecting to MAXIS
EMConnect ""

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
'----------------------------------------------------------------------------------------------------DAIL
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check = "DAIL" THEN
	EMReadScreen IEVS_type, 4, 6, 6 'read the DAIL msg'
	IF IEVS_type = "PARI" THEN
		EMSendKey "t"
		match_found = TRUE
		EMReadScreen MAXIS_case_number, 8, 5, 73
		MAXIS_case_number= TRIM(MAXIS_case_number)
		'----------------------------------------------------------------------------------------------------IEVP
	   'Navigating deeper into the match interface
	   CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
	   CALL write_value_and_transmit("INTM", 20, 71)   'navigates to IEVP
	   TRANSMIT
    END IF
ELSEIF dail_check <> "DAIL" THEN
 	CALL MAXIS_case_number_finder (MAXIS_case_number)
    MEMB_number = "01"
    BeginDialog case_number_dialog, 0, 0, 131, 65, "Case Number to clear match"
      EditBox 60, 5, 65, 15, MAXIS_case_number
      EditBox 60, 25, 30, 15, MEMB_number
      ButtonGroup ButtonPressed
        OkButton 20, 45, 50, 15
        CancelButton 75, 45, 50, 15
      Text 5, 30, 55, 10, "MEMB Number:"
      Text 5, 10, 50, 10, "Case Number:"
    EndDialog
	DO
		err_msg = ""
		Dialog case_number_dialog
		IF ButtonPressed = 0 THEN StopScript
			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
			If IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2 digit member number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
'----------------------------------------------------------------------------------------------------STAT
	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	EMwritescreen MEMB_number, 20, 76
	TRANSMIT
	EMReadscreen SSN_number_read, 11, 7, 42
	SSN_number_read = replace(SSN_number_read, " ", "")
	CALL navigate_to_MAXIS_screen("INFC" , "____")
	CALL write_value_and_transmit("INTM", 20, 71)
	CALL write_value_and_transmit(SSN_number_read, 3, 63)
END IF

EMReadScreen err_msg, 7, 24, 2
IF err_msg = "NO IEVS" THEN script_end_procedure_with_error_report("An error occurred in IEVP, please process manually.")'checking for error msg'

'----------------------------------------------------------------------------------------------------selecting the correct wage match
Row = 8
	DO
		EMReadScreen INTM_match_status, 2, row, 73 'DO loop to check status of case before we go into insm'
		'UR Unresolved, System Entered Only
		'PR Person Removed From Household
		'HM Household Moved Out Of State
		'RV Residency Verified, Person in MN
		'FR Failed Residency Verification Request
		'PC Person Closed, Not PARIS Interstate
		'CC Case Closed, Not PARIS Interstate
		EMReadScreen INTM_period, 5, row, 59
		IF INTM_match_status = "" THEN script_end_procedure_with_error_report("A pending PARIS match could not be found. The script will now end.")
		'IF INTM_match_status <> "RV" THEN
	    	INTM_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
        	"   " & INTM_period, vbYesNoCancel, "Please confirm this match")
			IF INTM_info_confirmation = vbNo THEN
            	row = row + 1
            	'msgbox "row: " & row
            	IF row = 18 THEN
                	PF8
				row = 8
				EMReadScreen INTM_match_status, 2, row, 73
				EMReadScreen INTM_period, 5, row, 59
            	END IF
        	END IF
		IF INTM_info_confirmation = vbYes THEN EXIT DO
    	IF INTM_info_confirmation = vbCancel THEN script_end_procedure_with_error_report("The script has ended. The match has not been acted on.")
	LOOP UNTIL INTM_info_confirmation = vbYes

'-----------------------------------------------------navigating into the match'
'msgbox "row: " & row
		CALL write_value_and_transmit("X", row, 3) 'navigating to insm'
		'Ensuring that the client has not already had a difference notice sent
		EMReadScreen notice_sent, 1, 8, 73
		EMReadScreen sent_date, 8, 9, 73
		sent_date = trim(sent_date)
		If trim(sent_date) <> "" then sent_date= replace(sent_date, " ", "/")
	'--------------------------------------------------------------------Client name
		EMReadScreen client_Name, 26, 5, 27
		client_name = trim(client_name)                         'trimming the client name
		IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This seperates the two names
			length = len(client_name)                           'establishing the length of the variable
			position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
			last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
			first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
		ELSEIF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
			length = len(first_name)                        	'trimming the 1st name
			position = InStr(first_name, " ")               	'establishing the length of the variable
			first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
		ELSE                                'In cases where the last name takes up the entire space, THEN the client name becomes the last name
			first_name = ""
			last_name = client_name
		END IF
		first_name = trim(first_name)
	'----------------------------------------------------------------------Minnesota active programs
	EMReadScreen MN_Active_Programs, 15, 6, 59
	MN_active_programs = Trim(MN_active_programs)
	MN_active_programs = Trim(MN_active_programs)
	MN_active_programs = replace(MN_active_programs, " ", ", ")

	'Month of the PARIS match
	EMReadScreen Match_Month, 5, 6, 27
	Match_month = replace(Match_Month, " ", "/")

	'--------------------------------------------------------------------PARIS match state & active programs-this will handle more than one state
	DIM state_array()
	ReDIM state_array(5, 0)
	add_state = 0

	Const row_num			= 1
	Const state_name		= 2
	Const match_case_num 	= 3
	Const contact_info		= 4
	Const progs 			= 5

	row = 13
	DO
	'-------------------------------------------------------Reading for each state active programs
		EMReadScreen state, 2, row, 3
		IF trim(state) = "" THEN
			EXIT DO
		ELSE
			'-------------------------------------------------------------------Case number for match state (if exists)
			EMReadScreen Match_State_Case_Number, 13, row, 9
			Match_State_Case_Number = trim(Match_State_Case_Number)
			IF Match_State_Case_Number = "" THEN Match_State_Case_Number = "N/A"
			Redim Preserve state_array(5, 	add_state)
			state_array(row_num, 			add_state) = row
			state_array(state_name, 		add_state) = state
			state_array(match_case_num, 	add_state) = Match_State_Case_Number
		'-------------------------------------------------------------------PARIS match contact information
			EMReadScreen phone_number, 23, row, 22
			phone_number = TRIM(phone_number)
			If phone_number = "Phone: (     )" then
				phone_number = ""
			Else
				EMReadScreen phone_number_ext, 8, row, 51
				phone_number_ext = trim(phone_number_ext)
				If phone_number_ext <> "" then phone_number = phone_number & " Ext: " & phone_number_ext
			End if
			'-------------------------------------------------------------------reading and cleaning up the fax number if it exists
			EMReadScreen fax_check, 8, row + 1, 37
			fax_check = trim(fax_check)
			If fax_check <> "" then
				EMReadScreen fax_number, 21, row + 1, 24
				fax_number = TRIM(fax_number)
			End if
			If fax_number = "Fax: (     )" then fax_number = ""
			Match_contact_info = phone_number & " " & fax_number
			state_array(contact_info, add_state) = Match_contact_info

			'-------------------------------------------------------------------trims excess spaces of Match_Active_Programs
	   		Match_Active_Programs = "" 'sometimes blanking over information will clear the value of the variable'
			'match_row = row           'establishing match row the same as the current state row. Needs another variables since we are only incrementing the match row in the loop. Row needs to stay the same for larger loop/next state.
			DO
				IF Match_Active_Programs = "" THEN EXIT DO
				EMReadScreen Match_Prog, 22, row, 60
	   			Match_Prog = TRIM(Match_Prog)
				IF Match_Prog = "FOOD SUPPORT" THEN  Match_Prog = "FS"
				IF Match_Prog = "HEALTH CARE" THEN Match_Prog = "HC"
	    		IF Match_Prog <> "" THEN Match_Active_Programs = Match_Active_Programs & Match_Prog & ", "
				row = row + 1
			LOOP
			Match_Active_Programs = trim(Match_Active_Programs)
			'takes the last comma off of Match_Active_Programs when autofilled into dialog if more more than one app date is found and additional app is selected
			IF right(Match_Active_Programs, 1) = "," THEN Match_Active_Programs = left(Match_Active_Programs, len(Match_Active_Programs) - 1)
			state_array(progs, add_state) = Match_Active_Programs
			row = state_array(row_num, add_state)		're-establish the value of row to read phone and fax info
			Match_contact_info = ""
			phone_number = ""
			fax_number = ""
			'-----------------------------------------------add_state allows for the next state to gather all the information for array'
			add_state = add_state + 1
			'MsgBox add_state
			row = row + 3
			IF row = 19 THEN
				EMReadScreen last_page_check, 21, 24, 2
				last_page_check = trim(last_page_check)
				IF last_page_check = ""  THEN
					PF8
					row = 13
				END IF
			END IF
		END IF
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

	'--------------------------------------------------------------------Dialog
	BeginDialog OP_Cleared_dialog, 0, 0, 361, 255, "PARIS Match Claim Entered"
	  EditBox 55, 5, 40, 15, MAXIS_case_number
	  EditBox 170, 5, 20, 15, MEMB_number
	  EditBox 260, 5, 45, 15, INTM_period
	  DropListBox 55, 25, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	  EditBox 170, 25, 20, 15, OT_resp_memb
	  EditBox 260, 25, 45, 15, discovery_date
	  DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS"+chr(9)+"SSI", OP_program
	  EditBox 130, 65, 30, 15, OP_from
	  EditBox 180, 65, 30, 15, OP_to
	  EditBox 245, 65, 35, 15, Claim_number
	  EditBox 305, 65, 45, 15, Claim_amount
	  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS"+chr(9)+"SSI", OP_program_II
	  EditBox 130, 85, 30, 15, OP_from_II
	  EditBox 180, 85, 30, 15, OP_to_II
	  EditBox 245, 85, 35, 15, Claim_number_II
	  EditBox 305, 85, 45, 15, Claim_amount_II
	  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS"+chr(9)+"SSI", OP_program_III
	  EditBox 130, 105, 30, 15, OP_from_III
	  EditBox 180, 105, 30, 15, OP_to_III
	  EditBox 245, 105, 35, 15, Claim_number_III
	  EditBox 305, 105, 45, 15, Claim_amount_III
	  DropListBox 85, 135, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Contact_other_state
	  DropListBox 265, 135, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", bene_other_state
	  DropListBox 185, 155, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
	  DropListBox 265, 155, 90, 15, "Agency Error"+chr(9)+"Household"+chr(9)+"Non-Collect--Agency Error"+chr(9)+"GRH Vendor"+chr(9)+"Fraud"+chr(9)+"Admit Fraud", collectible_reason
	  EditBox 50, 175, 45, 15, hc_claim_number
	  EditBox 185, 175, 40, 15, Fed_HC_AMT
	  EditBox 310, 175, 45, 15, HC_resp_memb
	  EditBox 50, 195, 150, 15, Reason_OP
	  EditBox 50, 215, 305, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 265, 235, 45, 15
	    CancelButton 310, 235, 45, 15
	  Text 5, 10, 50, 10, "Case Number: "
	  Text 130, 10, 30, 10, "MEMB #:"
	  Text 210, 10, 45, 10, "Match Period:"
	  Text 5, 30, 50, 10, "Fraud Referral:"
	  Text 110, 30, 55, 10, "OT Resp. Memb:"
	  Text 205, 30, 55, 10, "Discovery Date: "
	  GroupBox 5, 45, 350, 85, "Overpayment Information"
	  Text 15, 70, 30, 10, "Program:"
	  Text 105, 70, 20, 10, "From:"
	  Text 165, 70, 10, 10, "To:"
	  Text 215, 70, 25, 10, "Claim #"
	  Text 285, 70, 20, 10, "AMT:"
	  Text 285, 90, 20, 10, "AMT:"
	  Text 15, 110, 30, 10, "Program:"
	  Text 105, 110, 20, 10, "From:"
	  Text 165, 110, 10, 10, "To:"
	  Text 215, 110, 25, 10, "Claim #"
	  Text 285, 110, 20, 10, "AMT:"
	  Text 235, 160, 30, 10, "Reason:"
	  Text 145, 160, 40, 10, "Collectible?"
	  Text 150, 140, 115, 10, "Accessing benefits in other state?:"
	  Text 250, 180, 55, 10, "HC Resp. Memb:"
	  Text 5, 140, 75, 10, "Contacted other state?: "
	  Text 140, 180, 45, 10, "Fed HC AMT:"
	  Text 5, 200, 40, 10, "OP Reason:"
	  Text 5, 180, 40, 10, "HC Claim #:"
	  Text 105, 90, 20, 10, "From:"
	  Text 165, 90, 10, 10, "To:"
	  Text 215, 90, 25, 10, "Claim #"
	  Text 15, 90, 30, 10, "Program:"
	  Text 5, 220, 45, 10, "Other Notes:"
	  CheckBox 5, 160, 115, 10, "Out of state verification received", out_state_checkbox
	  CheckBox 235, 195, 120, 10, "Earned Income disregard allowed", EI_checkbox
	EndDialog

	Do
		err_msg = ""
		dialog OP_Cleared_dialog
		cancel_confirmation
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
		IF trim(Reason_OP) = "" or len(Reason_OP) < 8 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 8)."
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
		IF IEVS_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
		IF EI_allowed_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if Earned Income disregard was allowed."
	  	IF collectible_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if claim is collectible."
		IF collectible_dropdown = "YES" and collectible_reason = "" THEN err_msg = err_msg & vbnewline & "* Please advise why claim is collectible."
		IF contact_other_state = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if other state(s) have been contacted."
		IF bene_other_state = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if client received benefits in other state(s)."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)

	'Going to the MISC panel to add claim referral tracking information
	Call navigate_to_MAXIS_screen ("STAT", "MISC")
	Row = 6
    EmReadScreen panel_number, 1, 02, 73
	If panel_number = "0" then
		EMWriteScreen "NN", 20,79
		TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
		EmReadScreen MISC_error_msg,  74, 24, 02
		IF trim(MISC_error_msg) = "" THEN
	        case_note_only = FALSE
		ELSE
			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYesNo + vbQuestion, "Message handling")
			IF maxis_error_check = vbYes THEN
				case_note_only = TRUE 'this will case note only'
			END IF
			IF maxis_error_check= vbNo THEN
				case_note_only = FALSE 'this will update the panels and case note'
			END IF
		END IF
	ELSE
		IF case_note_only = FALSE THEN
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
    		If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
		END IF
		'writing in the action taken and date to the MISC panel
		EMWriteScreen "Claim Determination", Row, 30
		EMWriteScreen date, Row, 66
		PF3
	END IF 'checking to make sure maxis case is active'

    start_a_blank_case_note
    Call write_variable_in_case_note("-----Claim Referral Tracking-----")
	IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
    Call write_bullet_and_variable_in_case_note("Program(s)", programs)
	Call write_bullet_and_variable_in_case_note("Action Date", date)
	Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    Call write_variable_in_case_note("-----")
	Call write_variable_in_case_note(worker_signature)
	PF3

'-----------------------------------------------------------------------------------------CASENOTE
	start_a_blank_case_note
	CALL write_variable_in_CASE_NOTE ("-----" & INTM_period & " PARIS MATCH " & "(" & first_name &  ") OVERPAYMENT CLAIM ENTERED-----")
	Call write_bullet_and_variable_in_case_note("Client Name", Client_Name)
	Call write_bullet_and_variable_in_case_note("MN Active Programs", MN_active_programs)
	Call write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
	Call write_bullet_and_variable_in_case_note("Period", INTM_period)
	'formatting for multiple states
	FOR paris_match = 0 to Ubound(state_array, 2)
		CALL write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, paris_match) & " -----")
		Call write_bullet_and_variable_in_case_note("Match State Active Programs", state_array(progs, paris_match))
		Call write_bullet_and_variable_in_case_note("Match State Contact Info", state_array(contact_info, paris_match))
		Call write_bullet_and_variable_in_case_note("Match State Active Programs", state_array(progs, paris_match))
	NEXT
		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
		CALL write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
		IF OP_program_II <> "Select:" then CALL write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
		IF OP_program_III <> "Select:" then CALL write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
		IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
		Call write_bullet_and_variable_in_case_note("Client accessing benefits in other state", bene_other_state)
		Call write_bullet_and_variable_in_case_note("Contacted other state", contact_other_state)
		IF programs = "Health Care" or programs = "Medical Assistance" THEN
			Call write_bullet_and_variable_in_case_note("HC responsible members", HC_resp_memb)
			Call write_bullet_and_variable_in_case_note("HC claim number", hc_claim_number)
			Call write_bullet_and_variable_in_case_note("Total federal Health Care amount", Fed_HC_AMT)
			CALL write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
		END IF
		If out_state_checkbox = CHECKED THEN Call write_variable_in_case_note("Out of state verification received.")
		Call write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
		Call write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
		Call write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
		Call write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
		Call write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
		IF other_notes <> "" THEN Call write_bullet_and_variable_in_case_note("Other notes", other_notes)
		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
		CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
	PF3
	'gathering the case note for the email'
	IF programs = "Health Care" or programs = "Medical Assistance" THEN
		EMWriteScreen "x", 5, 3
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
		CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us","Claims entered for #" &  MAXIS_case_number & " Member # " & MEMB_number & " Date Overpayment Created: " & discovery_date & " Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)
	END IF
	'---------------------------------------------------------------writing the CCOL case note'
	'msgbox "Navigating to CCOL to add case note, please contact MiKayla with any concerns."
	'Call navigate_to_MAXIS_screen("CCOL", "CLSM")
	'EMWriteScreen Claim_number, 4, 9
	'Transmit
	'PF4
	'EMReadScreen existing_case_note, 1, 5, 6
	'IF existing_case_note = "" THEN
	'	PF4
	'ELSE
	'	PF9
	'END IF
	'	CALL write_variable_in_CCOL_NOTE ("-----" & INTM_period & " PARIS MATCH " & "(" & first_name &  ") OVERPAYMENT CLAIM ENTERED-----")
	'	CALL write_bullet_and_variable_in_CCOL_note("Client Name", Client_Name)
	'	CALL write_bullet_and_variable_in_CCOL_note("MN Active Programs", MN_active_programs)
	'	CALL write_bullet_and_variable_in_CCOL_note("Discovery date", discovery_date)
	'	CALL write_bullet_and_variable_in_CCOL_note("Period", INTM_period)
	'	write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
	'	'formatting for multiple states
	'	FOR paris_match = 0 to Ubound(state_array, 2)
	'		write_variable_in_CCOL_NOTE("----- Match State: " & state_array(state_name, paris_match) & " -----")
	'		CALL write_bullet_and_variable_in_CCOL_note("Match State Active Programs", state_array(progs, paris_match))
	'		CALL write_bullet_and_variable_in_CCOL_note("Match State Contact Info", state_array(contact_info, paris_match))
	'		CALL write_bullet_and_variable_in_CCOL_note("Match State Active Programs", state_array(progs, paris_match))
	'	NEXT
	'	write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
	'	write_variable_in_CCOL_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
	'	IF OP_program_II <> "Select:" then write_variable_in_CCOL_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
	'	IF OP_program_III <> "Select:" then write_variable_in_CCOL_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
	'	IF EI_checkbox = CHECKED THEN write_variable_in_CCOL_NOTE("* Earned Income Disregard Allowed")
	'	write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
	'	CALL write_bullet_and_variable_in_CCOL_note("Client accessing benefits in other state", bene_other_state)
	'	CALL write_bullet_and_variable_in_CCOL_note("Contacted other state", contact_other_state)
	'	IF programs = "Health Care" or programs = "Medical Assistance" THEN
	'		CALL write_bullet_and_variable_in_CCOL_note("HC responsible members", HC_resp_memb)
	'		CALL write_bullet_and_variable_in_CCOL_note("HC claim number", hc_claim_number)
	'		CALL write_bullet_and_variable_in_CCOL_note("Total federal Health Care amount", Fed_HC_AMT)
	'		write_variable_in_CCOL_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
	'	END IF
	'	If out_state_checkbox = CHECKED THEN Call write_variable_in_CCOL_note("Out of state verification received.")
	'	CALL write_bullet_and_variable_in_CCOL_note("Other responsible member(s)", OT_resp_memb)
	'	CALL write_bullet_and_variable_in_CCOL_note("Fraud referral made", fraud_referral)
	'	CALL write_bullet_and_variable_in_CCOL_note("Collectible claim", collectible_dropdown)
	'	CALL write_bullet_and_variable_in_CCOL_note("Reason that claim is collectible or not", collectible_reason)
	'	CALL write_bullet_and_variable_in_CCOL_note("Reason for overpayment", Reason_OP)
	'	IF other_notes <> "" THEN CALL write_bullet_and_variable_in_CCOL_note("Other notes", other_notes)
	'	write_variable_in_CCOL_NOTE("----- ----- ----- ----- ----- ----- -----")
	'	write_variable_in_CCOL_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
	'PF3 'exit the case note'
	'PF3 'back to dail'
script_end_procedure("Success PARIS match updated and please copy case note to CCOL.")
