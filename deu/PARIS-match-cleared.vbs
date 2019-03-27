name_of_script = "ACTIONS - DEU-PARIS MATCH CLEARED.vbs"
start_time = timer
STATS_counter = 1             'sets the stats counter at one
STATS_manualtime = 700        'manual run time in seconds
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

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/03/2018", "Updated coding for multiple states on INSM panel.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/28/2018", "Added handling for more than two states of PARIS matches on INSM.", "MiKayla Handley, Hennepin County")
CALL changelog_update("09/21/2018", "Added handling for more than one page of PARIS matches on INTM.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/18/2017", "Updates created to add options for sending difference notice and handling for resolution status and multiple states", "MiKayla Handley, Hennepin County")
CALL changelog_update("09/20/2017", "Updates made across the board, including action and case note", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/17/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------dialog
BeginDialog notice_action_dialog, 0, 0, 166, 85, "SEND DIFFERENCE NOTICE?"
  CheckBox 25, 35, 105, 10, "YES - Send Difference Notice", send_notice_checkbox
  CheckBox 25, 50, 130, 10, "NO - Continue Match Action to Clear", clear_action_checkbox
  ButtonGroup ButtonPressed
    OkButton 60, 65, 45, 15
    CancelButton 110, 65, 45, 15
  Text 10, 10, 145, 20, "A difference notice has not been sent, would you like to send the difference notice now?"
EndDialog

'---------------------------------------------------------------------THE SCRIPT
EMConnect ""

'warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
'If warning_box = vbCancel THEN stopscript

EMReadscreen dail_check, 4, 2, 48
IF dail_check <> "DAIL" THEN script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

EMReadScreen DAIL_message, 4, 6, 6 'read the DAIL msg'
IF DAIL_message <> "PARI" THEN script_end_procedure("This is not a Paris match. Please select a Paris match, and run the script again.")

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)
'msgbox "did i make it"
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   'navigates to INFC
CALL write_value_and_transmit("INTM", 20, 71)   'navigates to INTM
EMReadScreen error_msg, 2, 24, 2
error_msg = TRIM(error_msg)
IF error_msg <> "" THEN script_end_procedure("An error occured in INFC, please process manually.")'-------option to read from REPT need to checking for error msg'

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
				'msgbox "this should transmit"
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

IF notice_sent = "N" THEN
    DO
    	DO
    		err_msg = ""
    		Dialog notice_action_dialog
    		IF ButtonPressed = 0 THEN StopScript
    		IF send_notice_checkbox = UNCHECKED AND clear_action_checkbox = UNCHECKED THEN err_msg = err_msg & vbNewLine & "* Please select an answer to continue."
    		IF send_notice_checkbox = CHECKED AND clear_action_checkbox = CHECKED THEN err_msg = err_msg & vbNewLine & "* Please select only one answer to continue."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)
	LOOP UNTIL are_we_passworded_out = false
END IF

IF send_notice_checkbox = CHECKED THEN
'----------------------------------------------------------------Defaulting checkboxes to being CHECKED (per DEU instruction)
    diff_notice_CHECKBOX = 1
    shelter_verf_CHECKBOX = 1
    proof_residency_CHECKBOX = 1

    BeginDialog SEND_PARIS_DIFF_NOTICE_dialog, 0, 0, 376, 235, "SEND PARIS MATCH DIFFERENCE NOTICE"
    	Text 10, 15, 130, 10, "Case number: "   & MAXIS_case_number
    	Text 165, 15, 175, 10, "Client Name: "  & Client_Name
    	Text 10, 35, 110, 10, "Match month: "   & Match_Month
    	Text 165, 35, 175, 10, "MN active program(s): "   & MN_active_programs
	GroupBox 5, 50, 360, 75, "PARIS MATCH INFORMATION:"
		For item = 0 to Ubound(state_array, 2)
		    Text 10, 60, 75, 10, "Match State: "   & state_array(state_name, item)
		    Text 10, 75, 135, 10, "Match State Case Number: "   & state_array(match_case_num, item)
		    Text 10, 90, 155, 10, "Match State Active Programs:" & state_array(progs, item)
		    Text 10, 105, 360, 15, "Match State Contact Info: " & state_array(contact_info, item)
		Next
		   'For item = 1 to Ubound(state_array, 2)
			'   Text 185, 60, 110, 10, "2nd Match State: "   &  state_array(state_name, item)
			'   Text 185, 90, 185, 10, "2nd Match Active Programs: "   & state_array(progs, item)
			'   Text 185, 75, 110, 10, "2nd Match State Case Number: " & state_array(match_case_num, item)
			'   Text 185, 105, 175, 15, "2nd Match Contact Info: "  & state_array(contact_info, item)
		   'Next
    	Text 60, 180, 60, 10, "Referral to Fraud:"
      	Text 55, 160, 65, 10, "Contact Other State:"
    	Text 10, 140, 110, 10, "Accessing benefits in other state:"
      	DropListBox 120, 135, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", bene_other_state
      	DropListBox 120, 155, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", contact_other_state
      	DropListBox 120, 175, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"Undetermined", fraud_referral
    GroupBox 205, 130, 160, 50, "Verification requested:"
      CheckBox 210, 145, 50, 10, "Diff Notice", diff_notice_CHECKBOX
      CheckBox 290, 145, 70, 10, "Shelter Verification", shelter_verf_CHECKBOX
      CheckBox 210, 160, 70, 10, "Proof of Residency", proof_residency_CHECKBOX
      CheckBox 290, 160, 70, 10, "School Verification", schl_verf_CHECKBOX
      EditBox 120, 195, 245, 15, Other_Notes
      Text 75, 200, 40, 10, "Other notes:"
    ButtonGroup ButtonPressed
        OkButton 220, 215, 70, 15
        CancelButton 295, 215, 70, 15
    EndDialog

    	'---------------------------------------------------------------------send notice dialog and dialog DO...loop
    	DO
    		DO
    			err_msg = ""
    			Dialog SEND_PARIS_DIFF_NOTICE_dialog
    			IF ButtonPressed = 0 THEN StopScript
    			IF bene_other_state = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Is the client accessing benefits in other state?"
    			IF contact_other_state = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Did you contact the other state?"
				IF fraud_referral = "Select One:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
			LOOP UNTIL err_msg = ""
    		'--------------------------------------------------------------CHECKING FOR MAXIS WITHOUT TRANSMITTING SINCE THIS WILL NAVIGATE US AWAY FROM THE AREA WE ARE AT
    		EMReadScreen MAXIS_check, 5, 1, 39
    		IF MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " THEN
    			IF end_script = TRUE THEN
    				script_end_procedure("You Do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
    			ELSE
    				warning_box = MsgBox("You Do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
    				IF warning_box = vbCancel THEN stopscript
    			END IF
    		END IF
    	LOOP UNTIL MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

    	'sending the notice
    	PF9	'edit mode'
    	EMReadScreen edit_error, 2, 24, 2
    	edit_error = trim (edit_error)
    	IF edit_error <> "" THEN script_end_procedure ("Unable to send difference notice please review case")
    	EMwritescreen "Y", 8, 73 'send Notice
    	transmit

    	'--------------------------------------------------------------------The case note & case note related code
    	'creating new variable for case note for programs appealing that is incremential
    	pending_verifs = ""
    	IF shelter_verf_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Shelter, "
    	IF diff_notice_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
    	IF proof_residency_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Residency, "
    	IF schl_verf_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "School, "
    	'trims excess spaces of pending_verifs
    	pending_verifs = trim(pending_verifs)
    	'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
    	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
    	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

    	'-----------------------------------------------------------------------The case note
    	start_a_blank_CASE_NOTE
    	CALL write_variable_in_CASE_NOTE ("-----" & Match_month & " PARIS MATCH " & "(" & first_name &  ") DIFF NOTICE SENT-----")
    	CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
    	CALL write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
    	'formatting for multiple states
    	For item = 0 to Ubound(state_array, 2)
    		CALL write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, item) & " -----")
    		CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", state_array(progs, item))
    		CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", state_array(contact_info, item))
    	NEXT
    	CALL write_variable_in_CASE_NOTE ("-----")
    	CALL write_bullet_and_variable_in_CASE_NOTE("Client accessing benefits in other state", bene_other_state)
    	CALL write_bullet_and_variable_in_CASE_NOTE("Contacted other state", contact_other_state)
    	CALL write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
    	CALL write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
    	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    	CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
    	CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

ELSEIF clear_action_checkbox = CHECKED or notice_sent = "Y" THEN
	IF sent_date <> "" then MsgBox("A difference notice was sent on " & sent_date & ". The script will now navigate to clear the PARIS match.")
	BeginDialog PARIS_MATCH_CLEARED_dialog, 0, 0, 376, 260, "PARIS MATCH CLEARED"
     Text 10, 15, 130, 10, "Case number: "   & MAXIS_case_number
     Text 165, 15, 175, 10, "Client Name: "  & Client_Name
     Text 10, 35, 110, 10, "Match month: "   & Match_Month
     Text 165, 35, 175, 10, "MN active program(s): "   & MN_active_programs
	 GroupBox 5, 50, 360, 75, "PARIS MATCH INFORMATION:"
	 	For item = 0 to Ubound(state_array, 2)
     		Text 10, 60, 75, 10, "Match State: "   & state_array(state_name, item)
			Text 10, 75, 135, 10, "Match State Case Number: "   & state_array(match_case_num, item)
			Text 10, 90, 155, 10, "Match Active Programs: " & state_array(progs, item)
 			Text 10, 105, 170, 15, "Match State Contact Info: "   &  state_array(contact_info, item)
		Next
  	 Text 10, 140, 110, 10, "Accessing benefits in other state:"
     DropListBox 120, 135, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", bene_other_state
     DropListBox 120, 155, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", contact_other_state
     DropListBox 120, 175, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"Undetermined", fraud_referral
  	 GroupBox 205, 130, 160, 50, "Verification used to clear: "
	 CheckBox 210, 145, 50, 10, "Diff Notice", diff_notice_CHECKBOX
     CheckBox 290, 145, 70, 10, "Shelter Verification", shelter_verf_CHECKBOX
     CheckBox 210, 160, 70, 10, "Proof of Residency", proof_residency_CHECKBOX
     CheckBox 290, 160, 70, 10, "School Verification", schl_verf_CHECKBOX
     DropListBox 210, 195, 155, 15, "Select One:"+chr(9)+"PR - Person Removed From Household"+chr(9)+"HM - Household Moved Out Of State"+chr(9)+"RV - Residency Verified, Person in MN"+chr(9)+"FR - Failed Residency Verification Request"+chr(9)+"PC - Person Closed, Not PARIS Interstate"+chr(9)+"CC - Case Closed, Not PARIS Interstate", resolution_status
     EditBox 210, 220, 155, 15, Other_Notes
     Text 150, 200, 60, 10, "Resolution Status:"
     Text 170, 225, 40, 10, "Other notes:"
	 Text 60, 180, 60, 10, "Referral to Fraud:  "
	 Text 55, 160, 65, 10, "Contact other State:  "
  	 ButtonGroup ButtonPressed
     	OkButton 220, 240, 70, 15
     	CancelButton 295, 240, 70, 15
  	EndDialog

    DO
    	DO
    		err_msg = ""
    		Dialog PARIS_MATCH_CLEARED_dialog
    		IF ButtonPressed = 0 THEN StopScript
    		IF bene_other_state = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Is the client accessing benefits in other state?"
    		IF contact_other_state = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Did you contact the other state?"
			IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP UNTIL err_msg = ""

    	'CHECKING FOR MAXIS WITHOUT TRANSMITTING SINCE THIS WILL NAVIGATE US AWAY FROM THE AREA WE ARE AT
    	EMReadScreen MAXIS_check, 5, 1, 39
    	IF MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " THEN
    		IF end_script = True THEN
    			script_end_procedure("You Do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
    		ELSE
    			warning_box = MsgBox("You Do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
    			IF warning_box = vbCancel THEN stopscript
    		END IF
    	END IF
    LOOP UNTIL MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

    '--------------------------------------------------------------------The case note
    pending_verifs = ""
    IF shelter_verf_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Shelter, "
    IF Other_Verif_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other verification provided, "
    IF proof_residency_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "Residency, "
    IF schl_verf_CHECKBOX = CHECKED THEN pending_verifs = pending_verifs & "School, "
    'trims excess spaces of pending_verifs
    pending_verifs = trim(pending_verifs)
    'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
    IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

    Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
    'requested for HEADER of casenote'

    IF resolution_status = "PR - Person Removed From Household" THEN rez_status = "PR"
    IF resolution_status = "HM - Household Moved Out Of State" THEN rez_status = "HM"
    IF resolution_status = "RV - Residency Verified, Person in MN" THEN rez_status = "RV"
    IF resolution_status = "FR - Failed Residency Verification Request" THEN rez_status = "FR"
    IF resolution_status = "PC - Person Closed, Not PARIS Interstate" THEN rez_status = "PC"
    IF resolution_status = "CC - Case Closed, Not PARIS Interstate" THEN rez_status = "CC"

	'------------------------------------------------------------------'still need to be on PARIS Interstate Match Display (INSM)'
	PF9
	'msgbox rez_status
	EMwritescreen rez_status, 9, 27
	IF fraud_referral = "YES" THEN
		EMwritescreen "Y", 10, 27
		ELSE
		TRANSMIT
	END IF
	PF3
    PF3
	'CALL MAXIS_case_number
    '----------------------------------------------------------------the case match note
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE ("-----" & Match_month & " PARIS MATCH " & "(" & first_name &  ") CLEARED " & rez_status & "-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
    CALL write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
    'formatting for multiple states
    For item = 0 to Ubound(state_array, 2)
    	CALL write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, item) & " -----")
    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", state_array(progs, item))
    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", state_array(contact_info, item))
    NEXT
    CALL write_variable_in_CASE_NOTE ("-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Client accessing benefits in other state", bene_other_state)
    CALL write_bullet_and_variable_in_CASE_NOTE("Contacted other state", contact_other_state)
    CALL write_bullet_and_variable_in_CASE_NOTE("Verification used to clear", pending_verifs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Resolution Status", resolution_status)
	IF rez_status = "FR" THEN CALL write_variable_in_CASE_NOTE("Client has failed to cooperate with Paris Match - has not provided requested verifications showing they are living in MN. Client will need to provide this before the case is reopened ")
	CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
    CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
END IF

script_end_procedure("Success PARIS match updated.")
