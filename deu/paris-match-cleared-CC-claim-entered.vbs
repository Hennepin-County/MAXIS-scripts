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
call changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley") '#493
CALL changelog_update("10/20/2020", "Removed custom functions from script file. Functions have all been incorporated into the project's Function Library.", "Ilse Ferris, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/15/2019", "Updated script to copy case note to CCOL and clear matches at FR.", "MiKayla Handley, Hennepin County")
CALL changelog_update("09/28/2018", "Added handling for more than two states of PARIS matches on INSM.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updates made to correct error.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My DOcuments folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'---------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 131, 65, "Case Number to clear match"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 60, 25, 30, 15, MEMB_number
  ButtonGroup ButtonPressed
    OkButton 20, 45, 50, 15
    CancelButton 75, 45, 50, 15
  Text 5, 30, 55, 10, "MEMB Number:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog
DO
	    DO
        	err_msg = ""
        	Dialog Dialog1
        	cancel_without_confirmation
         	If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
         	If IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2 digit member number."
        	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        LOOP UNTIL err_msg = ""
		CALL check_for_password_without_transmit(are_we_passworded_out)
	Loop until are_we_passworded_out = false

''----------------------------------------------------------------------------------------------------STAT
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen MEMB_number, 20, 76
TRANSMIT
EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")

CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("INTM", 20, 71)
CALL write_value_and_transmit(SSN_number_read, 3, 63)
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

EMReadScreen error_check, 75, 24, 2
error_check = TRIM(error_check)
IF error_check <> "" THEN script_end_procedure(error_check & vbcr & "An error occurred, please process manually.") '-------option to read from REPT need to checking for error msg'
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
	IF trim(INTM_match_status) = "" THEN script_end_procedure_with_error_report("A pending PARIS match could not be found. The script will now end.")
	'IF INTM_match_status <> "RV" THEN
    	INTM_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
    	"   " & INTM_period, vbYesNoCancel, "Please confirm this match")
		IF INTM_info_confirmation = vbNo THEN
        	row = row + 1
			IF INTM_match_status = "" THEN script_end_procedure_with_error_report("A pending PARIS match could not be found. The script will now end.")
        	IF row = 18 THEN
            	PF8
				row = 8
				'EMReadScreen INTM_match_status, 2, row, 73
				'EMReadScreen INTM_period, 5, row, 59
			IF INTM_match_status = "" THEN script_end_procedure_with_error_report("A pending PARIS match could not be found. The script will now end.")
        	END IF
    	END IF
	IF INTM_info_confirmation = vbYes THEN EXIT DO
	IF INTM_info_confirmation = vbCancel THEN script_end_procedure_with_error_report("The script has ended. The match has not been acted on.")
LOOP UNTIL INTM_info_confirmation = vbYes
'-----------------------------------------------------navigating into the match'
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
client_name = UCase(client_name)
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
Const progs		     	= 5
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
		match_contact_info = phone_number & " " & fax_number
		state_array(contact_info, add_state) = match_contact_info
		'-------------------------------------------------------------------trims excess spaces of match_active_programs
   		match_active_programs = "" 'sometimes blanking over information will clear the value of the variable'
		match_row = row           'establishing match row the same as the current state row. Needs another variables since we are only incrementing the match row in the loop. Row needs to stay the same for larger loop/next state.
		DO
			EMReadScreen other_state_active_programs, 22, row, 60
   			other_state_active_programs = TRIM(other_state_active_programs)
			IF other_state_active_programs = "" THEN EXIT DO
			IF other_state_active_programs = "FOOD SUPPORT" THEN match_active_programs = match_active_programs & "FS, "
			IF other_state_active_programs = "HEALTH CARE" THEN match_active_programs = match_active_programs &  "HC, "
			IF other_state_active_programs = "STATE SSI" THEN match_active_programs = match_active_programs & "SSI, "
			IF other_state_active_programs = "NONE IDICATED" THEN match_active_programs = match_active_programs &  "NONE INDICATED"
			IF other_state_active_programs = "CASH" THEN match_active_programs = match_active_programs &  "CASH"
			IF other_state_active_programs = "CHILD CARE" THEN match_active_programs = match_active_programs &  "CCA"
			IF other_state_active_programs = "STATE WORKERS COMP" THEN match_active_programs = match_active_programs &  "WORKERS COMP"
    		row = row + 1
		LOOP
		match_active_programs = trim(match_active_programs)
		IF right(match_active_programs, 1) = "," THEN match_active_programs = left(match_active_programs, len(match_active_programs) - 1)
		state_array(progs, add_state) = match_active_programs
		row = state_array(row_num, add_state)		're-establish the value of row to read phone and fax info
		match_contact_info = ""
		phone_number = ""
		fax_number = ""
		'-----------------------------------------------add_state allows for the next state to gather all the information for array'
		add_state = add_state + 1
		'MsgBox add_state
		row = row + 3
		IF row = 19 THEN
			PF8
			EMReadScreen last_page_check, 21, 24, 2
			last_page_check = trim(last_page_check)
			IF last_page_check = ""  THEN row = 13
		END IF
	END IF
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"
'------------------------------------------------------------------'still need to be on PARIS Interstate Match Display (INSM)'
IF resolution_status = "PR - Person Removed From Household" THEN rez_status = "PR"
IF resolution_status = "HM - Household Moved Out Of State" THEN rez_status = "HM"
IF resolution_status = "RV - Residency Verified, Person in MN" THEN rez_status = "RV"
IF resolution_status = "FR - Failed Residency Verification Request" THEN rez_status = "FR"
IF resolution_status = "PC - Person Closed, Not PARIS Interstate" THEN rez_status = "PC"
IF resolution_status = "CC - Case Closed, Not PARIS Interstate" THEN rez_status = "CC"
PF9 'to edit the case'
'msgbox rez_status & "This update may cause issues, please let MiKayla know right away to assist with testing."
EMwritescreen rez_status, 9, 27
IF fraud_referral = "YES" THEN
	EMwritescreen "Y", 10, 27
	ELSE
	TRANSMIT
END IF
TRANSMIT
'--------------------------------------------------------------------Dialog
discovery_date = date
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 361, 290, "PARIS Match Claim Entered"
  EditBox 60, 5, 40, 15, MAXIS_case_number
  EditBox 200, 5, 20, 15, memb_number
  DropListBox 315, 5, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  EditBox 60, 25, 40, 15, discovery_date
  EditBox 200, 25, 20, 15, OT_resp_memb
  EditBox 315, 25, 40, 15, INTM_match_period
  DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
  EditBox 130, 65, 30, 15, OP_from
  EditBox 180, 65, 30, 15, OP_to
  EditBox 245, 65, 35, 15, Claim_number
  EditBox 305, 65, 45, 15, Claim_amount
  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
	  EditBox 130, 85, 30, 15, OP_from_II
	  EditBox 180, 85, 30, 15, OP_to_II
	  EditBox 245, 85, 35, 15, Claim_number_II
	  EditBox 305, 85, 45, 15, Claim_amount_II
	  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS"+chr(9)+"SSI", OP_program_III
	  EditBox 130, 105, 30, 15, OP_from_III
	  EditBox 180, 105, 30, 15, OP_to_III
	  EditBox 245, 105, 35, 15, Claim_number_III
	  EditBox 305, 105, 45, 15, Claim_amount_III
	  DropListBox 50, 125, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS"+chr(9)+"SSI", OP_program_IV
	  EditBox 130, 125, 30, 15, OP_from_IV
	  EditBox 180, 125, 30, 15, OP_to_IV
	  EditBox 245, 125, 35, 15, Claim_number_IV
	  EditBox 305, 125, 45, 15, Claim_amount_IV
	  EditBox 40, 155, 30, 15, HC_from
	  EditBox 90, 155, 30, 15, HC_to
	  EditBox 160, 155, 50, 15, HC_claim_number
	  EditBox 235, 155, 45, 15, HC_claim_amount
	  EditBox 100, 175, 20, 15, HC_resp_memb
	  EditBox 235, 175, 45, 15, Fed_HC_AMT
  CheckBox 10, 205, 50, 10, "Collectible?", collectible_checkbox
  DropListBox 100, 200, 100, 15, "Select:"+chr(9)+"Agency Error"+chr(9)+"Household"+chr(9)+"Non-Collect Agency Error"+chr(9)+"GRH Vendor"+chr(9)+"Fraud"+chr(9)+"Admit Fraud", collectible_reason
  CheckBox 10, 220, 120, 10, "Accessing benefits in other state?", bene_other_state_checkbox
  CheckBox 10, 235, 85, 10, "Contacted other state?", contact_other_state_checkbox
  CheckBox 230, 200, 120, 10, "Out of state verification received?", out_of_state_checkbox
  EditBox 305, 215, 45, 15, verif_rcvd_date
  CheckBox 230, 235, 125, 10, "Earned income disregard allowed?", EI_checkbox
  EditBox 50, 250, 305, 15, overpayment_reason
  DropListBox 75, 270, 140, 15, "Select One:"+chr(9)+"PR - Person Removed From Household"+chr(9)+"HM - Household Moved Out Of State"+chr(9)+"RV - Residency Verified, Person in MN"+chr(9)+"FR - Failed Residency Verification Request"+chr(9)+"PC - Person Closed, Not PARIS Interstate"+chr(9)+"CC - Case Closed, Not PARIS Interstate", resolution_status
  ButtonGroup ButtonPressed
    OkButton 260, 270, 45, 15
    CancelButton 310, 270, 45, 15
	GroupBox 5, 45, 350, 100, "Overpayment Information"
  Text 130, 55, 30, 10, "(MM/YY)"
  Text 180, 55, 30, 10, "(MM/YY)"
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
  Text 15, 90, 30, 10, "Program:"
  Text 15, 130, 30, 10, "Program:"
  Text 105, 130, 20, 10, "From:"
  Text 165, 130, 10, 10, "To:"
  Text 215, 130, 25, 10, "Claim #"
  Text 285, 130, 20, 10, "AMT:"
  Text 15, 180, 80, 10, "HC OT resp. Memb(s) #:"
  Text 160, 180, 75, 10, "Total federal HC AMT:"
  Text 10, 255, 40, 10, "OP reason:"
  Text 250, 220, 50, 10, "Date verif rcvd: "
  Text 215, 160, 20, 10, "AMT:"
  Text 15, 160, 20, 10, "From:"
  Text 130, 160, 25, 10, "Claim #"
  Text 75, 160, 10, 10, "To:"
  GroupBox 5, 145, 350, 50, "HC Programs Only"
  Text 265, 30, 45, 10, "Match period:"
  Text 135, 30, 60, 10, "OT resp. Memb #:"
  GroupBox 5, 45, 350, 100, "Overpayment Information"
  Text 5, 30, 55, 10, "Discovery date: "
  Text 165, 10, 30, 10, "Memb #:"
  Text 5, 10, 50, 10, "Case number: "
  Text 260, 10, 50, 10, "Fraud referral:"
  Text 70, 205, 30, 10, "Reason:"
  Text 10, 275, 65, 10, "Resolution Status:"
EndDialog
Do
	err_msg = ""
	dialog Dialog1
	cancel_without_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(overpayment_reason) = "" or len(overpayment_reason) < 8 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 8)."
	IF OP_program = "Select:" and HC_claim_number = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
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
	IF collectible_checkbox = CHECKED and collectible_reason = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise why claim is collectible."
	IF out_of_state_checkbox = CHECKED and verif_rcvd_date = "" THEN err_msg = err_msg & vbnewline & "* Please enter the date verification was received."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL check_for_password_without_transmit(are_we_passworded_out)
IF OP_program = "FS" or OP_program_II = "FS" or OP_program_III = "FS" or OP_program_IV = "FS" or OP_program = "MF" or OP_program_II = "MF" or OP_program_III = "MF" or OP_program_IV = "MF" THEN
	'Going to the MISC panel to add claim referral tracking information
	Call navigate_to_MAXIS_screen ("STAT", "MISC")
	Row = 6
	EmReadScreen panel_number, 1, 02, 73
	If panel_number = "0" then
		EMWriteScreen "NN", 20,79
		TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
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
	If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
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
	Call write_variable_in_case_note("-----")
	Call write_variable_in_case_note(worker_signature)
	PF3
END IF
'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("-----" & INTM_period & " PARIS MATCH " & "(" & first_name &  ") CLEARED " & rez_status & " CLAIM ENTERED-----")
Call write_bullet_and_variable_in_case_note("Client Name", Client_Name)
Call write_bullet_and_variable_in_case_note("MN Active Programs", MN_active_programs)
Call write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
Call write_bullet_and_variable_in_case_note("Period", INTM_period)
'formatting for multiple states
FOR paris_match = 0 to Ubound(state_array, 2)
	CALL write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, paris_match) & " -----")
	Call write_bullet_and_variable_in_case_note("Match State Active Programs", state_array(progs, paris_match))
	Call write_bullet_and_variable_in_case_note("Match State Contact Info", state_array(contact_info, paris_match))
NEXT
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then CALL write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then CALL write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF HC_claim_number <> "" THEN
	Call write_bullet_and_variable_in_case_note("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_case_note("HC claim number", hc_claim_number)
	Call write_bullet_and_variable_in_case_note("Total federal Health Care amount", Fed_HC_AMT)
	CALL write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF collectible_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Collectible claim")
IF collectible_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Non-Collectible claim")
IF collectible_reason <> "Select:" THEN Call write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
IF bene_other_state_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Client accessing benefits in other state")
IF contact_other_state_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Contacted other state")
If out_state_checkbox = CHECKED THEN Call write_variable_in_case_note("Out of state verification received.")
Call write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
Call write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
Call write_bullet_and_variable_in_case_note("Reason for overpayment", overpayment_reason)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3 'have to PF3 because more steps are taken'
'gathering the case note for the email'
IF HC_claim_number <> "" THEN
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
	CALL create_outlook_email("HSPH.FAA.Unit.AR.Spaulding@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & MEMB_number & " Date Overpayment Created: " & discovery_date & " Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)
END IF
'---------------------------------------------------------------writing the CCOL case note'
Call navigate_to_MAXIS_screen("CCOL", "CLSM")
EMWriteScreen Claim_number, 4, 9
TRANSMIT
EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
error_check = trim(error_check)
If error_check <> "" then script_end_procedure_with_error_report(error_check & "Unable to update this case. Please review case, and run the script again if applicable.")
PF4
EMReadScreen existing_case_note, 1, 5, 6
IF existing_case_note = "" THEN
	PF4
ELSE
	PF9
END IF
CALL write_variable_in_CCOL_note("-----" & INTM_period & " PARIS MATCH " & "(" & first_name &  ") OVERPAYMENT CLAIM ENTERED-----")
Call write_bullet_and_variable_in_CCOL_note("Client Name", Client_Name)
Call write_bullet_and_variable_in_CCOL_note("MN Active Programs", MN_active_programs)
Call write_bullet_and_variable_in_CCOL_note("Discovery date", discovery_date)
Call write_bullet_and_variable_in_CCOL_note("Period", INTM_period)
'formatting for multiple states
FOR paris_match = 0 to Ubound(state_array, 2)
	CALL write_variable_in_CCOL_note("----- Match State: " & state_array(state_name, paris_match) & " -----")
	Call write_bullet_and_variable_in_CCOL_note("Match State Active Programs", state_array(progs, paris_match))
	Call write_bullet_and_variable_in_CCOL_note("Match State Contact Info", state_array(contact_info, paris_match))
NEXT
CALL write_variable_in_CCOL_note("----- ----- ----- ----- -----")
CALL write_variable_in_CCOL_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then CALL write_variable_in_CCOL_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then CALL write_variable_in_CCOL_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF HC_claim_number <> "" THEN
	Call write_bullet_and_variable_in_CCOL_note("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CCOL_note("HC claim number", hc_claim_number)
	Call write_bullet_and_variable_in_CCOL_note("Total federal Health Care amount", Fed_HC_AMT)
	CALL write_variable_in_CCOL_note(" * Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Allowed")
IF collectible_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Collectible claim")
IF collectible_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note("* Non-Collectible claim")
IF collectible_reason <> "Select:" THEN Call write_bullet_and_variable_in_CCOL_note("Reason that claim is collectible or not", collectible_reason)
IF bene_other_state_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Client accessing benefits in other state")
IF contact_other_state_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Contacted other state")
If out_state_checkbox = CHECKED THEN Call write_variable_in_CCOL_note("Out of state verification received.")
Call write_bullet_and_variable_in_CCOL_note("Other responsible member(s)", OT_resp_memb)
Call write_bullet_and_variable_in_CCOL_note("Fraud referral made", fraud_referral)
Call write_bullet_and_variable_in_CCOL_note("Reason for overpayment", overpayment_reason)
CALL write_variable_in_CCOL_note("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CCOL_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3 'exit the case note because it is CCOL'
script_end_procedure_with_error_report("Success PARIS match overpayment case note entered and copied to CCOL please review case note to ensure accuracy.")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/24/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/24/2022
'--All variables in dialog match mandatory fields-------------------------------06/24/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------06/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/24/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------06/24/2022
'--PRIV Case handling reviewed -------------------------------------------------06/24/2022
'--Out-of-County handling reviewed----------------------------------------------06/24/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/24/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/24/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------06/24/2022
'--Script name reviewed---------------------------------------------------------06/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------------------------N/A
'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/24/2022
'--comment Code-----------------------------------------------------------------06/24/2022
'--Update Changelog for release/update------------------------------------------06/24/2022
'--Remove testing message boxes-------------------------------------------------06/24/2022
'--Remove testing code/unnecessary code-----------------------------------------06/24/2022
'--Review/update SharePoint instructions----------------------------------------06/24/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/24/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/24/2022
'--Complete misc. documentation (if applicable)---------------------------------06/24/2022
'--Update project team/issue contact (if applicable)----------------------------06/24/2022
