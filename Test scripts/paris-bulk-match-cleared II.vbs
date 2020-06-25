'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK PARIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("03/20/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'------------------------------------------------------------------------------------------------------establishing date variables

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates


'The dialog is defined in the loop as it can change as buttons are pressed
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 50, "Select the source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog
'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog
do
    dialog Dialog1
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

FOR i = 1 to 18	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets up the array to store all the information for each client'
Dim PARI_array()
ReDim PARI_array (8, 0)

DIM state_array()
ReDIM state_array(5, 0)
add_state = 0

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_number    	= 0			'Each of the case numbers will be stored at this position'
Const  memb_ss_num   	= 1
Const client_name_first	= 2
Const   ex_date			= 3
Const  	clear_code  	= 4
Const  state_date  		= 5
Const  ex_clear     	= 6
Const case_status  		= 7
Const case_notes   		= 8
'this is the state array '
Const row_num			= 9
Const state_name		= 10
Const match_case_num	= 11
Const contact_info		= 12
Const progs 			= 13

'Now the script adds all the clients on the excel list into an array
excel_row = excel_row_to_start 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	excel_state_match					= objExcel.cells(excel_row, 6).value	'establishes MAXIS case number
	MAXIS_case_number 					= objExcel.cells(excel_row, 7).value	'establishes MAXIS case number
    Client_SSN 							= objExcel.cells(excel_row, 8).value	'establishes client SSN
	first_name 							= objExcel.cells(excel_row, 10).value	'establishes client SSN
    excel_date_notice_sent				= objExcel.cells(excel_row, 11).value	'establishes if notice has already been sent
    contact_other_state                 = objExcel.cells(excel_row, 12).value	'establishes cleared status for the match
	worker_entered_cleared_status	    = objExcel.cells(excel_row, 13).value	'establishes cleared status for the match
	excel_date_cleared					= objExcel.cells(excel_row, 14).value
	'TODO other_notes so that workers can have a single note on a case
    'cleaned up
	excel_state_match					= trim(excel_state_match) 'remove extra spaces'
    MAXIS_case_number 					= trim(MAXIS_case_number)
    Client_SSN 							= trim(Client_SSN)
    Client_SSN 							= replace(Client_SSN, "-", "")
	first_name 							= trim(first_name)
    excel_date_notice_sent 				= trim(excel_date_notice_sent)
    worker_entered_cleared_status 	  	= trim(worker_entered_cleared_status)

	'Adding client information to the array'
	ReDim Preserve PARI_array(8,    entry_record)	'This resizes the array based on the number of rows in the Excel File'
	PARI_array (case_number,        entry_record) = MAXIS_case_number		'The client information is added to the array'
	PARI_array (memb_ss_num,     	entry_record) = Client_SSN
	PARI_array (client_name_first, 	entry_record) = first_name
    PARI_array (ex_date,        	entry_record) = excel_date_notice_sent
	PARI_array (clear_code,         entry_record) = worker_entered_cleared_status
	PARI_array (state_date,   		entry_record) = contact_other_state
    PARI_array (ex_clear,   		entry_record) = excel_date_cleared
	PARI_array (case_status, entry_record) = ""
	PARI_array (case_notes,  entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	Stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop

'msgbox entry_record

excel_row = excel_row_to_start
back_to_self


For i = 0 to Ubound(PARI_array, 2)
	'Establishing values for each case in the array of cases
	MAXIS_case_number	= PARI_array (case_number, i)
	Client_SSN 	    	= PARI_array (memb_ss_num, i)
    first_name          = PARI_array (client_name_first, i)

    msgbox MAXIS_case_number

	EMReadscreen dail_check, 4, 2, 48
	IF dail_check <> "DAIL" THEN script_end_procedure("You are not in your dail. This script will stop.")

	'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
	EMSendKey "T"
	TRANSMIT

	EMReadScreen DAIL_message, 4, 6, 6 'read the DAIL msg'
	IF DAIL_message <> "PARI" THEN script_end_procedure("This is not a Paris match. Please select a Paris match, and run the script again.")

	EMReadScreen MAXIS_case_number, 8, 5, 73
	MAXIS_case_number= TRIM(MAXIS_case_number)
	msgbox "did i make it"
	'Navigating deeper into the match interface
	CALL write_value_and_transmit("I", 6, 3)   'navigates to INFC
	CALL write_value_and_transmit("INTM", 20, 71)   'navigates to INTM
	EMReadScreen error_check, 75, 24, 2
	error_check = TRIM(error_check)
	IF error_check <> "" THEN script_end_procedure(error_check & vbcr & "An error occurred, please process manually.")'-------option to read from REPT need to checking for error msg'

	'-----------------------------------------------------navigating into the match'
	row = 8
	'Ensuring that match has not already been resolved.
	Do
		EMReadScreen INTM_match_status, 2, row, 73 'DO loop to check status of case before we go into insm'
		INTM_match_status = trim(INTM_match_status)
		'msgbox INTM_match_status
		EMReadScreen INTM_period, 5, row, 59
		If INTM_match_status = "" THEN
			objExcel.cells(excel_row, 17).value = "No pending match found. Please review INTM."
			case_note_actions = FALSE
			exit do
		ELSE
			If INTM_match_status = "UR" THEN
				case_note_actions = TRUE
				exit do
			End if
			row = row + 1
		END IF
	Loop until row = 17
	msgbox "row: " & row
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

	'converts the name to all CAPS
	client_name = UCase(client_name)
	'----------------------------------------------------------------------Minnesota active programs
	EMReadScreen MN_Active_Programs, 15, 6, 59
	MN_active_programs = Trim(MN_active_programs)
	MN_active_programs = Trim(MN_active_programs)
	MN_active_programs = replace(MN_active_programs, " ", ", ")

	'Month of the PARIS match
	EMReadScreen Match_Month, 5, 6, 27
	Match_month = replace(Match_Month, " ", "/")

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
			Redim Preserve state_array(9, 	add_state)
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
			'match_row = row           'establishing match row the same as the current state row. Needs another variables since we are only incrementing the match row in the loop. Row needs to stay the same for larger loop/next state.
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

	IF notice_sent = "N" and worker_entered_cleared_status = "" THEN
	PF9	'edit mode'
	MsgBox "we want to send a notice"
	EMReadScreen edit_error, 2, 24, 2
	edit_error = trim (edit_error)
	IF edit_error <> "" THEN script_end_procedure ("Unable to send difference notice please review case")
	EMwritescreen "Y", 8, 73 'send Notice
	TRANSMIT
	'--------------------------------------------------------------------The case note & case note related code
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
	pending_verifs = "Interstate Match Notice, Shelter Verification, Proof of Residency"

	'-----------------------------------------------------------------------The case note
	start_a_blank_CASE_NOTE
	CALL write_variable_in_CASE_NOTE ("-----" & Match_month & " PARIS MATCH " & "(" & first_name &  ") DIFF NOTICE SENT-----")
	CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
	CALL write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
	'formatting for multiple states
	For i = 0 to Ubound(state_array, 2)
		CALL write_variable_in_CASE_NOTE("----- Match State: " & state_array(state_name, i) & " -----")
		CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", state_array(progs, i))
		CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", state_array(contact_info, i))
	NEXT
	CALL write_variable_in_CASE_NOTE ("-----")
	CALL write_bullet_and_variable_in_CASE_NOTE("Client accessing benefits in other state", bene_other_state)
	CALL write_bullet_and_variable_in_CASE_NOTE("Contacted other state", contact_other_state)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
	CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

	'ensuring that the case note saved. If not, adding it to the notes for the user to review.
	PF3
	EMReadScreen note_date, 8, 5, 6
	If note_date <> current_date then
		PARI_array(case_status, i) = "Error"
		PARI_array(case_notes, i) = "Case note does not appear to have been saved."	'Explanation for the rejected report'
	   Else
        PARI_array(case_status, i) = "Case updated"
        PARI_array(case_notes, i) = ""	'Explanation for the rejected report'
	End if


    ObjExcel.Cells(Excel_row, 11).Value = PARI_array(sent_date,   i)
    ObjExcel.Cells(Excel_row, 16).Value = PARI_array(case_status, i)
    ObjExcel.Cells(Excel_row, 17).Value = PARI_array(case_notes,  i)

	Excel_row = Excel_row + 1

    match_active_programs = ""
	match_contact_info = ""
	phone_number = ""
	fax_number = ""
	Msgbox match_programs & " test"
Next

FOR i = 1 to 19	'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Stats_counter = stats_counter + 1
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
