'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - METS RETRO HEALTH CARE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 230          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/27/2023", "Added MS Word output of case note to copy/paste in METS case note.", "Ilse Ferris, Hennepin County")
call changelog_update("03/07/2022", "Updated team emails from 601 to team 603 for retro processing. Added FIAT checkbox for retro determination option.", "Ilse Ferris, Hennepin County")
call changelog_update("05/21/2021", "Updated browser to default when opening SIR from Internet Explorer to Edge.", "Ilse Ferris, Hennepin County")
call changelog_update("10/20/2020", "Updated link to REQUEST TO APPL use form on SharePoint.", "Ilse Ferris, Hennepin County")
call changelog_update("03/04/2019", "Per project request - Removed checkbox for DOA and scenario pushbuttons.", "MiKayla Handley, Hennepin County")
call changelog_update("11/04/2019", "Updated link in SharePoint to Request to APPL useform.", "Ilse Ferris, Hennepin County")
call changelog_update("11/04/2019", "Updated link to Useform and changed name of form to 'Request to APPL'.", "Ilse Ferris, Hennepin County")
call changelog_update("07/26/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'-------------------------------------------------------------------------------------------------------script
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 146, 90, "METS Retro Health Care"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 60, 30, 55, 15, METS_case_number
  DropListBox 60, 50, 80, 15, "Select One:"+chr(9)+"Initial Request"+chr(9)+"Proofs Received"+chr(9)+"Retro Determination", initial_option
  ButtonGroup ButtonPressed
    OkButton 60, 70, 40, 15
    CancelButton 100, 70, 40, 15
  Text 5, 55, 50, 10, "Select process:"
  Text 10, 15, 50, 10, "MAXIS case #:"
  Text 10, 35, 50, 10, "METS case #:"
EndDialog

'Main dialog: user will input case number and member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1
		cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
		call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(METS_case_number) = False or len(METS_case_number) <> 8 then err_msg = err_msg & vbNewLine & "* Enter a valid METS case number."
        If initial_option = "Select One:" then err_msg = err_msg & vbcr & "* Select a retro process."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
If PRIV_check = "PRIV" then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member information - putting into an array
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
    EMReadscreen ref_nbr, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    'EMReadScreen client_DOB, 10, 8, 42
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

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BEGINDIALOG Dialog1, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
    Text 10, 5, 105, 10, "Household members to look at:"
    FOR i = 0 to total_clients											'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
    NEXT
    ButtonGroup ButtonPressed
    OkButton 185, 10, 50, 15
    CancelButton 185, 30, 50, 15
ENDDIALOG

Do
    Dialog Dialog1       'runs the dialog that has been dynamically created. Streamlined with new functions.
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

check_for_maxis(False)

HC_membs = -1
Dim HC_array()
ReDim HC_array(6, 0)

const member_number_const   = 0
const member_name_const     = 1
const retro_months_const    = 2
const retro_scenario_const  = 3
const retro_status_const    = 4
const retro_reason_const    = 5

FOR i = 0 to total_clients
    IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
        HC_membs = HC_membs + 1
        ReDim Preserve HC_array(6, HC_membs)
        HC_array(member_number_const, HC_membs) = left(all_clients_array(i, 0), 2)
        HC_array(member_name_const, HC_membs) = all_clients_array(i, 0)
        HC_array(retro_months_const, HC_membs) = ""
        HC_array(retro_scenario_const, HC_membs) = ""
        HC_array(retro_status_const, HC_membs) = ""
        HC_array(retro_reason_const, HC_membs) = ""
    END IF
NEXT

If initial_option = "Initial Request" then
    due_date = dateadd("d", 10, date) & ""
    case_note_header = "METS Initial Retro Request"

    ''-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 311, (190 + (HC_membs * 15)), "Initial Retro Request for #" & MAXIS_case_number
    Text 5, 15, 80, 10, "METS Application Date:"
    EditBox 85, 10, 50, 15, application_date
    Text 180, 15, 70, 10, "Verification Due Date:"
    EditBox 255, 10, 50, 15, due_date
    Text 10, 40, 70, 10, "Verifs/Forms Needed:"
    ComboBox 85, 35, 220, 15, "SELECT OR TYPE"+chr(9)+"Client verbally attested - No changes"+chr(9)+"DHS-3960 - Request for Information"+chr(9)+"Income for requested retro months", forms_needed
    Text 40, 60, 45, 10, "Other Notes:"
    EditBox 85, 55, 220, 15, other_notes
    CheckBox 10, 80, 140, 10, "METS DOA is 11 months prior to today.", DOA_checkbox
    CheckBox 10, 95, 185, 10, "Request to APPL use form created/sent.", useform_checkbox
    CheckBox 195, 95, 110, 10, "Task created for retro request.", task_checkbox
    ButtonGroup ButtonPressed
    Text 10, 125, 35, 10, "Scenario"
    x = 0
    For item = 0 to Ubound(HC_array, 2)
        If HC_array(member_name_const, item) <> "" then
            DropListBox 10, (140 + (x * 15)), 50, 15, "Select:"+chr(9)+"A"+chr(9)+"B"+chr(9)+"C"+chr(9)+"D"+chr(9)+"E", HC_array(retro_scenario_const, item)
            DropListBox 65, (140 + (x * 15)), 50, 15, "Select:"+chr(9)+"1 Month"+chr(9)+"2 Months"+chr(9)+"3 Months", HC_array(retro_months_const, item)
            Text 130, (145 + (x * 15)), 165, 10, HC_array(member_name_const, item)
            x = x + 1
        End if
    Next
    GroupBox 5, 110 , 300, (50 + (x * 15)), "For each applicant, enter the retro months and scenario:"
    Text 10, (175 + (x * 15)), 60, 10, "Worker Signature:"
    EditBox 75, (170 + (x * 15)), 130, 15, worker_signature
    ButtonGroup ButtonPressed
    OkButton 210, (170 + (x * 15)), 45, 15
    CancelButton 260, (170 + (x * 15)), 45, 15
    Text 70, 125, 45, 10, "Retro months"
    Text 135, 125, 60, 10, "Applicant's Name"
    EndDialog

    'Main dialog: user will input case number and member number
    DO
        DO
            'establishing value of variable, this is necessary for the Do...LOOP
            err_msg = ""
            Dialog Dialog1
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            If isdate(application_date) = false  then err_msg = err_msg & vbcr & "* Enter a valid application date."
            If forms_needed <> "Client verbally attested - No changes" then
                If isdate(due_date) = false then err_msg = err_msg & vbcr & "* Enter a valid verificaiton due date."
            End if
            IF trim(forms_needed) = "SELECT OR TYPE" then err_msg = err_msg & vbcr & "* Enter or select the verifications and/or forms needed."
            'retro scenario
            For item = 0 to ubound(HC_array, 2)
            	If (HC_array(retro_scenario_const, item)) = "Select:" and hc_array(member_name_const, item) <> "" then err_msg = err_msg & vbCr & "* Select a retro scenario for " & hc_array(member_name_const, item)
            Next
            'amt of retro months
            For item = 0 to ubound(HC_array, 2)
                If (HC_array(retro_months_const, item)) = "Select:" and hc_array(member_name_const, item) <> "" then err_msg = err_msg & vbCr & "* Select amount of retro months requested for " & hc_array(member_name_const, item)
            NEXT
            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End if

If initial_option = "Proofs Received" then
    case_note_header = "METS Retro Proofs Received"
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 296, (125 + (HC_membs * 15)), "Retro Proofs Received for #" & MAXIS_case_number
    Text 10, 20, 30, 10, "Scenario"
    x = 0
    For item = 0 to Ubound(HC_array, 2)
        If HC_array(member_name_const, item) <> "" then
            DropListBox 10, (35 + (x * 15)), 50, 15, "Select:"+chr(9)+"A"+chr(9)+"B"+chr(9)+"C"+chr(9)+"D"+chr(9)+"E", HC_array(retro_scenario_const, item)
            Text 75, (35 + (x * 15)), 200, 10, HC_array(member_name_const, item)
            x = x + 1
        End if
    Next
    GroupBox 5, 5, 285, (50 + (x * 15)), "Enter the retro scenario for each applicant:"
    CheckBox 05, (60 + (x * 15)), 140, 10, "All verifications and/or forms received.", verifs_checkbox
    Text 5, (85+ (x * 15)), 45, 10, "Other Notes:"
    EditBox 55, (80 + (x * 15)), 235, 15, other_notes
    Text 5, (110 + (x * 15)), 60, 10, "Worker Signature:"
    EditBox 70, (105 + (x * 15)), 130, 15, worker_signature
    ButtonGroup ButtonPressed
    OkButton 205, (105 + (x * 15)), 40, 15
    CancelButton 250, (105 + (x * 15)), 40, 15
    Text 75, 20, 60, 10, "Applicant's Name"
    EndDialog

    'Main dialog: user will input case number and member number
    DO
        DO
            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
            Dialog Dialog1
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            For item = 0 to Ubound(HC_array, 2)
                If (HC_array(retro_scenario_const, item)) = "Select:" and hc_array(member_name_const, item) <> "" then err_msg = err_msg & vbCr & "* Select a retro scenario for " & hc_array(member_name_const, item)
            NEXT
            If verifs_checkbox = 0 then err_msg = err_msg & vbNewLine & "* All verifications and forms must be receieved."
            IF worker_signature = "" then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End if

If initial_option = "Retro Determination" then
    case_note_header = "METS Retro Determination"
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 356, (130 + (HC_membs * 15)), "Retro Determination for #" & MAXIS_case_number
    x = 0
    For item = 0 to Ubound(HC_array, 2)
        If HC_array(member_name_const, item) <> "" then
            DropListBox 10, (30 + (x * 15)), 50, 15, "Select:"+chr(9)+"Approved"+chr(9)+"Denied", HC_array(retro_status_const, item)
            Editbox 75, (30 + (x * 15)), 180, 15, HC_array(retro_reason_const, item)
            Text 270, (35 + (x * 15)), 130, 10, HC_array(member_name_const, item)
            x = x + 1
        End if
    Next
    GroupBox 5, 5, 345, (50 + (x * 15)), "For each applicant, enter the retro determination:"
    Text 10, (65 + (x * 15)), 45, 10, "Other Notes:"
    EditBox 70, (60 + (x * 15)), 280, 15, other_notes
    CheckBox 70, (80 + (x * 15)), 140, 10, "HC eligibility was FIATed.", fiat_checkbox
    Text 5, (100 + (x * 15)), 60, 10, "Worker Signature:"
    EditBox 70, (95 + (x * 15)), 150, 15, worker_signature
    ButtonGroup ButtonPressed
    OkButton 265, (95 + (x * 15)), 40, 15
    CancelButton 310, (95 + (x * 15)), 40, 15
    Text 75, 20, 105, 10, "App'd months or Denial Reason"
    Text 15, 20, 45, 10, "Determination"
    Text 270, 20, 60, 10, "Applicant's Name"
    EndDialog

    'Main dialog: user will input case number and member number
    DO
        DO
            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
            Dialog Dialog1
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            'retro determination
            For item = 0 to ubound(HC_array, 2)
                If (HC_array(retro_status_const, item)) = "Select:" and hc_array(member_name_const, item) <> "" then err_msg = err_msg & vbCr & "* Select a retro determination for " & hc_array(member_name_const, item)
            NEXT
            'approved months or denial reason
            For item = 0 to ubound(HC_array, 2)
                If trim((HC_array(retro_reason_const, item))) = "" and hc_array(member_name_const, item) <> "" then err_msg = err_msg & vbCr & "* Enter approved months or denial reason(s) for " & hc_array(member_name_const, item)
            NEXT
            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End if

'Creating the member numbers output
member_numbers = ""
For i = 0 to ubound(hc_array, 2)
    IF hc_array(member_name_const, i) <> "" then
        member_numbers = member_numbers & hc_array(member_number_const, i) & ", "
    End if
Next

'Entering member numbers in the header of case note
member_numbers = trim(member_numbers) 'trims excess spaces of member_numbers
If right(member_numbers, 1) = "," THEN member_numbers = left(member_numbers, len(member_numbers) - 1) 'takes the last comma off of member_numbers

memb_info = " for Memb " & member_numbers ' for the case note

'----------------------------------------------------------------------------------------------------Outlook's time to shine
email_content = ""

If DOA_checkbox = checked then
    send_email = True
    email_content = "* METS DOA (Date of Application) was 11 months prior to today." & vbcr & vbcr
    team_email = "603"
elseif initial_option = "Retro Determination" then
    send_email = False
    team_email = ""
else
    If HC_array(retro_scenario_const, 0) = "B"  or HC_array(retro_scenario_const, 0) = "C" or HC_array(retro_scenario_const, 0) = "D" or HC_array(retro_scenario_const, 0) = "E" then
        send_email = True
        team_email = "603"
    Elseif HC_array(retro_scenario_const, 0) = "A" then
        send_email = False
        team_email = ""
    End if
End if

For i = 0 to ubound(HC_array, 2)
    If HC_array(member_name_const, i) <> "" then
        If initial_option = "Initial Request" then
            household_info = household_info & vbcr & " - " & HC_array(member_name_const, i) & ": Scenario " & HC_array(retro_scenario_const, i) & ". Requested " & HC_array(retro_months_const, i) & " retro coverage."
        Elseif initial_option = "Proofs Received" then
            household_info = household_info & vbcr & " - " & HC_array(member_name_const, i) & ": Scenario " & HC_array(retro_scenario_const, i)
        elseif initial_option = "Retro Determination" then
            household_info = household_info & vbcr & " - " & HC_array(member_name_const, i) & ": Retro coverage " & HC_array(retro_status_const, i) & "--" & HC_array(retro_reason_const, i)
        End if
    End if
Next

additional_content = ""
If trim(other_notes) <> "" then additional_content = additional_content & vbcr & "Other Notes: " & other_notes
If trim(forms_needed) <> "" then additional_content = additional_content & vbcr & "Verifs/forms needed: " & forms_needed
If verifs_checkbox = 1 then additional_content = additional_content & vbcr & "* All verifications and/or forms received."

email_header = initial_option & " for " & MAXIS_case_number & " - Action Required"
body_of_email = email_content & "---Health Care Member Information---" & household_info & vbcr & additional_content

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
IF send_email = True THEN CALL create_outlook_email("HSPH.EWS.Team." & team_email, "", email_header, body_of_email, "", True)

'------------------------------------------------------------------------------------Case Note
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(case_note_header & memb_info)
Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", METS_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Date of Application", application_date)
'doesn't add due date if no changes are reported. This case can be worked on now.
If forms_needed <> "Client verbally attested - No changes" and trim(due_date) <> "" then Call write_bullet_and_variable_in_CASE_NOTE("Verfication Due Date", due_date)
Call write_bullet_and_variable_in_CASE_NOTE("Verifs/Forms Requested", forms_needed)
Call write_variable_in_CASE_NOTE("---Health Care Member Information---")
'HH member array output
For i = 0 to ubound(HC_array, 2)
    If HC_array(member_name_const, i) <> "" then
        If initial_option = "Initial Request" then
            Call write_variable_in_CASE_NOTE(" - " & HC_array(member_name_const, i) & ": Scenario " & HC_array(retro_scenario_const, i) & ". Requested " & HC_array(retro_months_const, i) & " retro coverage.")
        Elseif initial_option = "Proofs Received" then
            Call write_variable_in_CASE_NOTE(" - " & HC_array(member_name_const, i) & ": Scenario " & HC_array(retro_scenario_const, i))
        elseif initial_option = "Retro Determination" then
            Call write_variable_in_CASE_NOTE(" - " & HC_array(member_name_const, i) & ": Retro coverage " & HC_array(retro_status_const, i) & "--" & HC_array(retro_reason_const, i))
        End if
    End if
Next

IF DOA_checkbox = 1 then Call write_variable_in_case_note("* Team " & team_email & " emailed re: DOA over 11 months prior to current date.")
IF useform_checkbox = 1 then CALL write_variable_in_case_note("* Request to APPL use form created and sent.")
If task_checkbox = 1 then Call write_variable_in_case_note("* Task created in METS for retro request.")
If verifs_checkbox = 1 then Call write_variable_in_case_note("* All verification and/or forms received for retro determination.")
If fiat_checkbox = 1 then Call write_variable_in_case_note("* HC eligibility was fiated.")
If send_email = True then Call write_variable_in_case_note("* Email notification sent to " & team_email & ".")
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

'----------------------------------------------------------------------------------------------------For METS CASE Note
'The METS team is also supposed to create case notes. The following code will read the case note, and replace and copy it to a MS Word document with exemption below (line 425)
message_array = ""
Call write_value_and_transmit("X", 5, 3)
note_row = 4			'Beginning of the case notes
Do 						'Read each line
    EMReadScreen note_line, 76, note_row, 3
    note_line = trim(note_line)
    If instr(note_line, "METS Case Number") then note_line = "* MAXIS Case Number: " & MAXIS_case_number 'replaces METS Case Number with MAXIS Case number for the METS folks
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

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
objWord.Caption = "METS Case Note"
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 30
objSelection.PageSetup.BottomMargin = 25
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "14"
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.TypeText "METS Case Note Verbiage - Copy/Paste into METS if needed. If not, close this document w/o saving:" & vbcr & vbcr
objSelection.TypeText message_array

If initial_option = "Initial Request" then
    navigate_decision = Msgbox("Do you want to open a Request to APPL useform?", vbQuestion + vbYesNo, "Navigate to Useform?")
    If navigate_decision = vbYes then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://aem.hennepin.us/rest/services/HennepinCounty/Processes/ServletRenderForm:1.0?formName=HSPH5004_1-0.xdp&interactive=1"
    If navigate_decision = vbNo then navigate_to_form = False
End if

If send_email = True then
    script_end_procedure_with_error_report("An email notification was sent to " & team_email & "." & vbcr & "A Word document has been created to copy/paste the MAXIS case note into METS case notes if applicable.")
else
    script_end_procedure_with_error_report("Success, your case note has been created." & vbcr & "A Word document has been created to copy/paste the MAXIS case note into METS case notes if applicable.")
End if

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/07/2022
'--Tab orders reviewed & confirmed----------------------------------------------02/27/2023
'--Mandatory fields all present & Reviewed--------------------------------------03/07/2022
'--All variables in dialog match mandatory fields-------------------------------03/07/2022
'Review dialog names for content and content fit in dialog----------------------02/27/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------03/07/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------03/07/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------02/27/2023---------------Removed to read the note for the MS Word output
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-02/27/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/07/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------03/07/2022-----------------N/A
'--PRIV Case handling reviewed -------------------------------------------------03/07/2022
'--Out-of-County handling reviewed----------------------------------------------03/07/2022-----------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/07/2022
'--BULK - review output of statistics and run time/count (if applicable)--------03/07/2022-----------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------02/27/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------03/07/2022
'--Incrementors reviewed (if necessary)-----------------------------------------03/07/2022-----------------N/A
'--Denomination reviewed -------------------------------------------------------03/07/2022
'--Script name reviewed---------------------------------------------------------03/07/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------03/07/2022-----------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------03/07/2022
'--comment Code-----------------------------------------------------------------03/07/2022
'--Update Changelog for release/update------------------------------------------02/27/2023
'--Remove testing message boxes-------------------------------------------------03/07/2022
'--Remove testing code/unnecessary code-----------------------------------------03/07/2022
'--Review/update SharePoint instructions----------------------------------------02/27/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------03/07/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------03/07/2022
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------02/27/2023
'--Complete misc. documentation (if applicable)---------------------------------03/07/2022-----------------N/A
'--Update project team/issue contact (if applicable)----------------------------03/07/2022
