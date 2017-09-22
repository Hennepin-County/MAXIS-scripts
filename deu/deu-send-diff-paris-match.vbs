'GATHERING STATS===========================================================================================
name_of_script = "DEU–SEND DIFF PARIS NOTICE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 0
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/13/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

EMReadScreen DAIL_message, 4, 6, 6 'read the DAIL msg'
If DAIL_message <> "PARI" then script_end_procedure("This is not a Paris match. Please select a Paris match, and run the script again.")

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)

'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   'navigates to INFC
CALL write_value_and_transmit("INTM", 20, 71)   'navigates to INTM
EMReadScreen error_msg, 2, 24, 2
error_msg = TRIM(error_msg)
If error_msg <> "" then script_end_procedure("An error occured in INFC, please process manually.")'checking for error msg'

Row = 8
Do
	EMReadScreen Status, 2, row, 73 'do loop to check status of case before we go into insm'
	IF Status <> "UR" then
		row = row + 1
    ELSE
		exit do
	End if
Loop until trim(Status) = "" or row = 19

CALL write_value_and_transmit("X", row, 3) 'navigating to insm'

'Ensuring that the client has not already had a difference notice sent
EMReadScreen notice_sent, 1, 8, 73
EMReadScreen sent_date, 8, 9, 73
sent_date = replace(sent_date, " ", "/")
If notice_sent = "Y" then script_end_procedure("A difference notice was already sent 'on " & sent_date & ". The script will now end.")

'Formatting the client name for the spreadsheet
'----------------------------------------------------------------------------------------------------Client name
'Reading client name and splitting out the 1st name
EMReadScreen Client_Name, 26, 5, 27
'Formatting the client name for the spreadsheet
client_name = trim(client_name)                         'trimming the client name
if instr(client_name, ",") then    						'Most cases have both last name and 1st name. This seperates the two names
	length = len(client_name)                           'establishing the length of the variable
	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
elseif instr(first_name, " ") then   						'If there is a middle initial in the first name, then it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
Else                                'In cases where the last name takes up the entire space, then the client name becomes the last name
	first_name = ""
	last_name = client_name
END IF

'Minnesota active programs
EMReadScreen MN_Active_Programs, 15, 6, 59
MN_active_programs = Trim(MN_active_programs)
MN_active_programs = Trim(MN_active_programs)
MN_active_programs = replace(MN_active_programs, " ", ", ")

'Month of the PARIS match
EMReadScreen Match_Month, 5, 6, 27
Match_month = replace(Match_Month, " ", "/")

'Active programs in PARIS match state - this will handle more than one state
DIM state_array()
ReDIM state_array(5, 0)
add_state = 0

Const row_num			= 1
Const state_name		= 2
Const match_case_num 	= 3
Const contact_info		= 4
Const progs 			= 5 
		
row = 13
Do 
	'PARIS match state & active programs
	EMReadScreen state, 2, row, 3 
	If trim(state) = "" then 
		exit do 
	Else 
		'Case number for match state (if exists)
		EMReadScreen Match_State_Case_Number, 13, row, 9
		Match_State_Case_Number = trim(Match_State_Case_Number)
		If Match_State_Case_Number = "" then Match_State_Case_Number = "N/A"
		Redim Preserve state_array(5, 	add_state)
		state_array(row_num, 			add_state) = row
		state_array(state_name, 		add_state) = state
		state_array(match_case_num, 	add_state) = Match_State_Case_Number
		item = item + 1
	End if 
	row = row + 3
	If row = 19 then	
		PF8
		EMReadScreen last_page_check, 21, 24, 2
	End if 
Loop until last_page_check = "THIS IS THE LAST PAGE"

For item = 0 to Ubound(state_array, 2)
	row = state_array(row_num, item)
    Match_Active_Programs = ""
    Do 
    	Do
    		EMReadScreen Match_Prog, 22, row, 60
    		Match_Prog = TRIM(Match_Prog)
    		IF Match_Prog <> "" then Match_Active_Programs = Match_Active_Programs & Match_Prog & ", " 
			row = row + 1 
    	Loop until Match_Prog = "" or row = 19
    	If row = 19 then 
    		PF8 
    		EMReadScreen last_page_check, 21, 24, 02
    		If last_page_check <> "THIS IS THE LAST PAGE" then row = 13	're-establishes row for the new page
    	End if
    LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"
	
	'trims excess spaces of Match_Active_Programs
	Match_Active_Programs = trim(Match_Active_Programs)
	'takes the last comma off of Match_Active_Programs when autofilled into dialog if more more than one app date is found and additional app is selected
	If right(Match_Active_Programs, 1) = "," THEN Match_Active_Programs = left(Match_Active_Programs, len(Match_Active_Programs) - 1)	
	state_array(progs, item) = Match_Active_Programs 
	
	'PARIS match contact information
	EMReadScreen Phone_Number, 23, 13, 22
	Phone_Number = TRIM(Phone_Number)
	EMReadScreen Phone_Number_ext, 8, 13, 51
	Phone_Number_ext = trim(Phone_Number_ext)
	If Phone_Number_ext <> "" then Phone_Number = Phone_Number & " Ext " & Phone_Number_ext
	
	'establishing variable for PARIS match state contact information (with phone number and fax if applicable)
	Match_contact_info = ""
	'reading and cleaning up the fax number if it exists
	EMReadScreen fax_check, 8, 13, 37
	fax_check = trim(fax_check)
	If fax_check <> "" then
		EMReadScreen fax_number_number, 21, 14, 24
		fax_number = TRIM(fax_number)
	End if
	
	 If fax_number = "" then
		Match_contact_info = Match_contact_info & Phone_Number
	Else
	 	Match_contact_info = Match_contact_info & Phone_Number & ", " & fax_number
	End if
	state_array(contact_info, item) = Match_contact_info
next 

'Defaulting checkboxes to being checked (per DEU instruction)
Diff_Notice_Checkbox = 1
Shelter_Verf_CheckBox = 1
Proof_Residency_checkbox = 1

BeginDialog SEND_PARIS_DIFF_NOTICE_dialog, 0, 0, 376, (240 + (add_state * 80)), "Paris Match Send Difference Notice"
  Text 35, 15, 85, 10, "Case number: "  & MAXIS_case_number
  Text 125, 15, 235, 10, "Client Name: " & Client_Name
  Text 35, 35, 70, 10, "Match month: "  & Match_Month
  Text 125, 35, 235, 10, "MN active program(s): " & MN_active_programs  
  'handling for mutliple states
  For item = 0 to Ubound(state_array, 2)
   GroupBox 5, (55 + (add_state * 55)), 360, 70, "PARIS Match found in " & 		state_array(state_name, 	item)
   Text 10, 75, 	(225 + (add_state * 80)), 10, "Match State Case Number: " & state_array(match_case_num, item)
   Text 10, 90, 	(265 + (add_state * 80)), 10, "Match Active Programs: " & 	state_array(progs, 		 	item)
   Text 10, 105, 	(265 + (add_state * 80)), 10, "Match contact info: " & 		state_array(contact_info, 	item)
  Next
  Text 10, 140, 		   (130 + (add_state * 80)), 10, "Client accessing benefits in other state:"
  DropListBox 140, 135, 	(55 + (add_state * 80)), 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", bene_other_state
  Text 70, 165, 		 	(65 + (add_state * 80)), 10, "Contact Other State:  & Contact_other_state"
  DropListBox 140, 160,  	(55 + (add_state * 80)), 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", Contact_other_state
  GroupBox 205, 130, 	   (160 + (add_state * 80)), 70, "Verification requested:"
  CheckBox 210, 150,	 	(50 + (add_state * 80)), 10, "Diff Notice", Diff_Notice_Checkbox
  CheckBox 290, 150, 	 	(80 + (add_state * 80)), 10, "Shelter Verification", Shelter_Verf_CheckBox
  CheckBox 210, 165, 	 	(70 + (add_state * 80)), 10, "Proof of Residency", Proof_Residency_checkbox
  CheckBox 290, 165, 	 	(75 + (add_state * 80)), 10, "School Verification", Schl_verf_checkbox
  Text 5, 200, 			 	(45 + (add_state * 80)), 10, "Other notes:"
  EditBox 50, 195, 		   (200 + (add_state * 80)), 15, Other_Notes
  ButtonGroup ButtonPressed
    OkButton 260, 195, 	 	(50 + (add_state * 80)), 15
    CancelButton 315, 195, 	(50 + (add_state * 80)), 15
EndDialog

'dialog and dialog DO...loop
Do
	Do
		err_msg = ""
		Dialog SEND_PARIS_DIFF_NOTICE_dialog
		If ButtonPressed = 0 then StopScript
		If bene_other_state = "Select One:" then err_msg = err_msg & vbNewLine & "* Is the client accessing benefits in other state?"
		If Contact_other_state = "Select One:" then err_msg = err_msg & vbNewLine & "* Did you contact the other state?"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	
	'CHECKING FOR MAXIS WITHOUT TRANSMITTING SINCE THIS WILL NAVIGATE US AWAY FROM THE AREA WE ARE AT
	EMReadScreen MAXIS_check, 5, 1, 39
	If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
		If end_script = True then
			script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
		Else
			warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
			If warning_box = vbCancel then stopscript
		End if
	End if
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'sending the notice
PF9	'edit mode'
EMReadScreen edit_error, 2, 24, 2
edit_error = trim (edit_error)
if edit_error <> "" then script_end_procedure ("Unable to send difference notice please review case")
EMwritescreen "Y", 8, 73 'send Notice
transmit

'----------------------------------------------------------------------------------------------------The case note & case note related code
'creating new variable for case note for programs appealing that is incremential
pending_verifs = ""
If Shelter_Verf_CheckBox = checked THEN pending_verifs = pending_verifs & "Shelter, "
If Diff_Notice_Checkbox = checked then pending_verifs = pending_verifs & "Difference Notice, "
If Proof_Residency_checkbox = checked then pending_verifs = pending_verifs & "Residency, "
If Schl_verf_checkbox = checked then pending_verifs = pending_verifs & "School, "
'trims excess spaces of pending_verifs
pending_verifs = trim(pending_verifs)
'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

'The case note
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE (Match_month & " PARIS MATCH (" & first_name & ") DIFF NOTICE SENT")
Call write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
Call write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
'formatting for multiple states
For item = 0 to Ubound(state_array, 2)
	Call write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, item) & " -----")
	Call write_bullet_and_variable_in_CASE_NOTE("Match Active Programs", state_array(progs, item))
	Call write_bullet_and_variable_in_CASE_NOTE("Match Contact Info", state_array(contact_info, item))	
Next 
Call write_variable_in_CASE_NOTE ("*")
Call write_bullet_and_variable_in_CASE_NOTE("Client accessing benefits in other state", bene_other_state)
Call write_bullet_and_variable_in_CASE_NOTE("Contacted other state", Contact_other_state)
Call write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
Call write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

script_end_procedure("")
