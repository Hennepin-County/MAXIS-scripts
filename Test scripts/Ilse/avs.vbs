'GATHERING STATS===========================================================================================
name_of_script = "NOTES - AVS.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denominatinon = "I"
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
call changelog_update("10/21/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'----------------------------------------------------------------------------------------------------Initial dialog 
BeginDialog Initial_dialog, 0, 0, 151, 75, "AVS Initial Process Dialog"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  DropListBox 60, 35, 85, 15, "Select one..."+chr(9)+"AVS Form Processing"+chr(9)+"AVS Submission", initial_option
  ButtonGroup ButtonPressed
    OkButton 60, 55, 40, 15
    CancelButton 105, 55, 40, 15
  Text 10, 15, 45, 10, "Case number:"
  Text 5, 40, 50, 10, "Select process:"
EndDialog

'Main dialog: user will input case number and member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog initial_dialog				    
		cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
		call validate_MAXIS_case_number(err_msg, "*")
        If initial_option = "Select one..." then err_msg = err_msg & vbcr & "* Select an AVS process option."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member information 
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")

	client_string = ref_nbr & last_name & first_name
	client_array = client_array & trim(client_string) & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

DIM all_client_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	all_clients_array(x, 1) = 1    '1 = checked
NEXT

BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "Member Selection Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
    Text 5, 5, 180, 10, "Select members who are REQUIRED to sign AVS forms:"
	FOR i = 0 to total_clients											'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
    Text 5, (20 + (i * 15)), 50, 10, "Select option:"
    If initial_option = "AVS Form Processing" then DropListBox 60, (20 + (i * 15)), 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received – Incomplete", avs_option 
    If initial_option = "AVS Submission" then DropListBox 60, (20 + (i * 15)), 90, 15, "Select one..."+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_option 
	ButtonGroup ButtonPressed
	OkButton 185, 10, 50, 15
	CancelButton 185, 30, 50, 15
ENDDIALOG

Do 
    Do 
        err_msg = ""
        Dialog HH_memb_dialog       'runs the dialog that has been dynamically created. Streamlined with new functions.
        cancel_without_confirmation
        If avs_option = "Select one..." then err_msg = err_msg & vbcr & "* Select an AVS process option."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
check_for_maxis(False)

avs_membs = -1
Dim avs_members_array()
ReDim avs_members_array(2, 0)

const member_number_const   = 0
const member_name_const     = 1  
const hc_type_const         = 2

FOR i = 0 to total_clients
	IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
		avs_membs = avs_membs + 1
        ReDim Preserve avs_members_array(2, avs_membs)
        avs_members_array(member_number_const, avs_membs) = left(all_clients_array(i, 0), 2)
        avs_members_array(member_name_const, avs_membs) = all_clients_array(i, 0)
        avs_members_array(hc_type_const, avs_membs) = ""
	END IF
NEXT

'----------------------------------------------------------------------------------------------------AVS Forms processing option
If initial_option = "AVS Form Processing" then
    Select case avs_option
    
    Case "Initial Request"
        date_text = "Sent"
        set_TIKL = True 
    Case "Not Received"
        date_text = "Sent"
        set_TIKL = False 
    Case "Received - Complete"
        date_text = "Rec'd"
        set_TIKL = True 
    Case  "Received – Incomplete"
        date_text = "Rec'd"
        set_TIKL = True 
    End Select     

    BeginDialog non_MAGI_dialog, 0, 0, 271, (150 + (avs_membs * 15)), "Non-MAGI Referral for #" & MAXIS_case_number
      Text 10, 10, 55, 10, "Date of Request:"
      EditBox 70, 5, 55, 15, request_date
      Text 140, 10, 70, 10, "METS Case Number:"
      EditBox 210, 5, 55, 15, METS_case_number
      Text 5, 30, 65, 10, "Service requested:"
      DropListBox 70, 25, 195, 15, "Select one..."+chr(9)+"21+ years, no dependents and Medicare or SSI"+chr(9)+"65 years old, no dependents"+chr(9)+"Certified disabled, applying for MA-EPD"+chr(9)+"Certified disabled, Requesting waiver"+chr(9)+"Certified disabled, Requesting TEFRA"+chr(9)+"Child in Foster Care"+chr(9)+"Only Medicare Savings Programs requested"+chr(9)+"Other", service_requested
      CheckBox 5, 45, 70, 10, "SMRT approved.", SMRT_approved
      CheckBox 100, 45, 70, 10, "SMRT pending.", SMRT_pending
      CheckBox 5, 60, 180, 10, "Case has known duplicate PMI's and/or PMI issues.", PMI_checkbox
      Text 5, 80, 40, 10, "Other notes:"
      EditBox 50, 75, 215, 15, other_notes
      x = 0
      FOR item = 0 to ubound(avs_members_array, 2)							'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
          Text 10, (110 + (x * 15)), 140, 10, avs_members_array(member_name_const, item)
          If avs_members_array(member_name_const, item) <> "" then DropListBox 160, (110 + (x * 15)), 50, 15, "Select one..."+chr(9)+"MA"+chr(9)+"MCRE"+chr(9)+"IA"+chr(9)+"QHP", avs_members_array(hc_type_const, item)
          x = x + 1
      NEXT
      GroupBox 5, 95, 260, (20 + (x * 12)), "Client Information and current METS coverage:"
      Text 5,  (125 + (x * 12)), 60, 10, "Worker Signature:" 
      EditBox 70, (120 + (x * 12)), 85, 15, worker_signature
      ButtonGroup ButtonPressed
      OkButton 160, (120 + (x * 12)), 50, 15
      CancelButton 215, (120 + (x * 12)), 50, 15
    EndDialog
    
    'Main dialog: user will input case number and member number
    DO
        DO
            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
            Dialog non_MAGI_dialog				    
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            If isdate(request_date) = false or trim(request_date) = "" then err_msg = err_msg & vbcr & "* Enter a valid request date."
            If trim(METS_case_number) = "" or IsNumeric(METS_case_number) = False or len(METS_case_number) <> 8 then err_msg = err_msg & vbcr & "* Enter a valid METS case number."
            IF service_requested = "Select one..." then err_msg = err_msg & vbcr & "* Enter the service request reason."
            'HC_type
            For item = 0 to ubound(avs_members_array, 2)	
            	If (avs_members_array(hc_type_const, item)) = "Select one..."then err_msg = err_msg & vbCr & "* Select a health care type for each member."
            NEXT
            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
elseif initial_option = "AVS Submission" then 
    '(initial_option = "2. Request to end eligibility in METS" or "3. Eligibility ended in METS")
    BeginDialog elig_ended_dialog, 0, 0, 271, (120 + (avs_membs * 15)), "Eligibility Ending for #" & MAXIS_case_number
    If initial_option = "3. Eligibility ended in METS" then 
        Text 10, 10, 70, 10, "MMIS elig end date:"
        EditBox 80, 5, 55, 15, mmis_end_date
    End if 
      Text 140, 10, 70, 10, "METS Case Number:"
      EditBox 210, 5, 55, 15, METS_case_number
      x = 0
      FOR item = 0 to ubound(avs_members_array, 2)							'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
          Text 10, (40 + (x * 15)), 100, 10, avs_members_array(member_name_const, item)
          x = x + 1
      NEXT
      GroupBox 5, 25, 260, (25 + (x * 10)), "Client(s) name"
      Text 5, (65 + (x * 10)), 40, 10, "Other notes:"
      EditBox 50, (60 + (x * 10)), 215, 15, other_notes
      Text 5, (85 + (x * 10)), 60, 10, "Worker Signature:"
      EditBox 70, (80 + (x * 10)), 85, 15, worker_signature
      ButtonGroup ButtonPressed
          OkButton 160, (80 + (x * 10)), 50, 15
          CancelButton 215, (80 + (x * 10)), 50, 15
    EndDialog

    DO
        DO
            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
            dialog elig_ended_dialog				    
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            If (initial_option = "3. Eligibility ended in METS" AND (isdate(mmis_end_date) = false or trim(mmis_end_date) = "")) then err_msg = err_msg & vbcr & "* Enter a valid MMIS end date."
            If trim(METS_case_number) = "" or IsNumeric(METS_case_number) = False or len(METS_case_number) <> 8 then err_msg = err_msg & vbcr & "* Enter a valid METS case number."
            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End if 

client_name_list = ""
member_numbers = ""
For i = 0 to ubound(avs_members_array, 2)
    IF avs_members_array(member_name_const, i) <> "" then
        member_numbers = member_numbers & avs_members_array(member_number_const, i) & ", "
        'splitting up the client name to get the 1st name 
        client_name = avs_members_array(member_name_const, i)
        client_name = right(client_name, len(client_name) - 3)
        If instr(client_name, " ") then    						'Most cases have both last name and 1st name. This seperates the two names
            length = len(client_name)                           'establishing the length of the variable
            position = InStr(client_name, " ")                  'sets the position at the deliminator (in this case the comma)    
            first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
        END IF
        'adding first name to name list 
        first_name = trim(first_name)
        Call fix_case(first_name, 0)
        client_name_list = client_name_list & first_name & ", "    
    End if 
Next 

member_numbers = trim(member_numbers) 'trims excess spaces of member_numbers
If right(member_numbers, 1) = "," THEN member_numbers = left(member_numbers, len(member_numbers) - 1) 'takes the last comma off of member_numbers

client_name_list = trim(client_name_list) 'trims excess spaces of client_name_list
If right(client_name_list, 1) = "," THEN client_name_list = left(client_name_list, len(client_name_list) - 1) 'takes the last comma off of client_name_list

memb_info = " for Memb " & member_numbers ' for the case note

'----------------------------------------------------------------------------------------------------The case note
'Headers for case note 
If initial_option = "1. Non-MAGI referral" then header = "MA NON MAGI Referral"
If initial_option = "2. Request to end eligibility in METS" then header = "Requested METS eligibility to end"
If initial_option = "3. Eligibility ended in METS" then header = "Eligibility ended in METS effective " & mmis_end_date
If initial_option = "MAXIS to METS Migration" then header = "Closed HC " & CM_plus_1_mo & "/" & CM_plus_1_yr

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(header & memb_info) 
Call write_bullet_and_variable_in_CASE_NOTE("METS case number", METS_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Date of request", request_date)
Call write_variable_in_CASE_NOTE("---Health Care Member Information---") 
'HH member array output 
For i = 0 to ubound(avs_members_array, 2)
    If avs_members_array(member_name_const, i) <> "" then 
        If initial_option = "1. Non-MAGI referral" then  
            Call write_variable_in_CASE_NOTE(" - " & avs_members_array(member_name_const, i) & ", Current METS coverage: " & avs_members_array(hc_type_const, i))
        else 
            Call write_variable_in_CASE_NOTE(" - " & avs_members_array(member_name_const, i))
        End if    
    End if 
Next 
Call write_bullet_and_variable_in_CASE_NOTE("Service requested", service_requested)
Call write_bullet_and_variable_in_CASE_NOTE("MMIS eligibility end date", mmis_end_date)
If SMRT_approved = 1 then Call write_variable_in_CASE_NOTE("* SMRT is approved.")
If SMRT_pending = 1 then Call write_variable_in_CASE_NOTE("* SMRT is pending.")
If PMI_checkbox = 1  then Call write_variable_in_CASE_NOTE("* Case has known duplicate PMI/PMI issues.")
'METS to MAXIS case note only 
If initial_option = "MAXIS to METS Migration" then 
    Call write_variable_in_CASE_NOTE("* This case was identified by DHS as requiring conversion to the METS system.")
    If METS_case_number = "" then 
        Call write_variable_in_CASE_NOTE("* No associated METS case exists for the listed members.")
        Call write_variable_in_CASE_NOTE("* Informational notice generated via SPEC/MEMO to client regarding applying through mnsure.org.")
    Else 
        'For cases with affliated METS cases 
        Call write_variable_in_CASE_NOTE("* Informational notice generated via SPEC/MEMO to client. The METS team will contact the client if any additional information is needed to make a determination.")
    End if 
End if 

Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes) 
Call write_variable_in_CASE_NOTE("---") 
Call write_variable_in_CASE_NOTE(worker_signature) 


CALL create_maxis_friendly_date(date, 10, 5, 18)
Call write_variable_in_TIKL(variable)

script_end_procedure("")