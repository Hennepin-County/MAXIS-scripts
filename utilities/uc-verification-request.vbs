'GATHERING STATS===========================================================================================
name_of_script = "UTILITIES - UC VERIFICATION REQUEST.vbs"
start_time = timer
STATS_counter = 0
STATS_manualtime = 300
STATS_denominatinon = "M"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
function custom_HH_member_custom_dialog(HH_member_array)
'--- This function creates an array of all household members in a MAXIS case, and allows users to select which members to seek/add information to add to edit boxes in dialogs.
'~~~~~ HH_member_array: should be HH_member_array for function to work
'===== Keywords: MAXIS, member, array, dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
    transmit

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
        EMReadScreen access_denied_check, 13, 24, 2
        'MsgBox access_denied_check
        If access_denied_check = "ACCESS DENIED" Then
            PF10
            EMWaitReady 0, 0
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
        Else
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
    		last_name = trim(replace(last_name, "_", "")) & " "
    		first_name = trim(replace(first_name, "_", "")) & " "
    		mid_initial = replace(mid_initial, "_", "")
        End If
		client_string = ref_nbr & last_name & first_name & mid_initial
		client_array = client_array & client_string & "|"
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
		all_clients_array(x, 1) = 0         'unchecking all the members in this HH
	NEXT

	BEGINDIALOG HH_memb_dialog, 0, 0, 241, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		Text 10, 5, 105, 10, "Household members to look at:"
		FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
			IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		NEXT
		ButtonGroup ButtonPressed
		OkButton 185, 10, 50, 15
		CancelButton 185, 30, 50, 15
	ENDDIALOG
													'runs the dialog that has been dynamically created. Streamlined with new functions.
	Dialog HH_memb_dialog
	Cancel_without_confirmation
	check_for_maxis(True)

	HH_member_array = ""

	FOR i = 0 to total_clients
		IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
			IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
				'msgbox all_clients_
				HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
			END IF
		END IF
	NEXT

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
	HH_member_array = SPLIT(HH_member_array, " ")
end function
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("09/29/2023", "Improved handling for HH member dialog, updated dialog options to be more consistent", "Mark Riegel, Hennepin County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("06/23/2022", "Update to uncheck all HH members ", "MiKayla Handley, Hennepin County") '#882
call changelog_update("05/19/2022", "Update to send all UC verification requests to DEED email at: HSPH.ES.DEED@Hennepin.us", "Ilse Ferris, Hennepin County") '#847
call changelog_update("11/24/2021", "Updates to the dialog as most teams are now using ES support staff for UC verification request.", "MiKayla Handley") '#644'
call changelog_update("07/30/2021", "Inital Version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Grabs the case number
EMConnect ""
EMReadScreen mx_row_1, 79, 1, 1
EMReadScreen mx_row_2, 79, 2, 1
EMReadScreen mx_row_3, 79, 3, 1
EMReadScreen mx_row_4, 79, 4, 1
script_run_lowdown = "START OF SCRIPT" & vbCr & mx_row_1 & vbCr & mx_row_2 & vbCr & mx_row_3 & vbCr & mx_row_4
script_run_lowdown = script_run_lowdown & vbCr & "=============================================================" & vbCr & vbCr

CALL check_for_MAXIS(False)
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "Request for Unemployment Insurance Verification email has been sent." 'setting up closing_message UCriable for possible additions later based on conditions
'----------------------------------------------------------------------------------------------------Initial dialog
initial_help_text = "*** Unemployment Compensation ***" & vbNewLine & vbNewLine & "For residents under the age of 18 please see" & vbNewLine & "CM0010.18.01 - MANDATORY VERIFICATIONS - CASH" & vbNewLine & "CM0010.18.02 - MANDATORY VERIFICATIONS - SNAP" & vbNewLine & "Unemployment Insurance benefits are considered countable unearned income for all programs." & vbNewLine & vbNewLine &"NOTE: UC(Unemployment Compensation) and UI(Unemployment Income) are used interchangeably by workers."

'---------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 301, 95, "Verification Request for Unemployment Compensation"
  EditBox 85, 5, 50, 15, MAXIS_case_number
  CheckBox 10, 55, 25, 10, "CCA", cca_checkbox
  CheckBox 45, 55, 85, 10, "Other (please specify):", other_checkbox
  EditBox 130, 50, 160, 15, other_check_editbox
  ButtonGroup ButtonPressed
    OkButton 195, 75, 45, 15
    CancelButton 245, 75, 45, 15
    PushButton 165, 5, 15, 15, "!", initial_help_button
    PushButton 185, 5, 105, 15, "Unemployment Compensation", HSR_manual_button
  Text 35, 10, 50, 10, "Case Number:"
  GroupBox 5, 40, 290, 30, "Department (if outside ES)"
  Text 5, 25, 295, 10, "FYI: Overpaypayment, tax, child/spousal support deductions will be reviewed for this case."
EndDialog

DO
    DO
        err_msg = ""
        DO
            Dialog Dialog1
            cancel_without_confirmation
            IF ButtonPressed = initial_help_button then tips_tricks_msg = MsgBox(initial_help_text, vbInformation, "Tips and Tricks") 'see initial_help_text above for details of the text
            IF buttonpressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Unemployment_Insurance.aspx") 'HSR manual policy page
        LOOP until ButtonPressed = -1
        Call validate_MAXIS_case_number(err_msg, "*")
        IF cca_checkbox + other_checkbox = 2 Then err_msg = err_msg & vbCr & "* You cannot check both the 'CCA' and 'Other' boxes. Check only one option if selecting a department outside of ES."
        IF other_checkbox = CHECKED and trim(other_check_editbox) = "" THEN err_msg = err_msg & vbCr & "* You must fill out the field to specify which department you are in."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & "Resolve for the following for the script to continue." & vbcr & err_msg
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to  assword back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Determinging MAXIS Region and giving the user the option to send script if in inquiry
Call back_to_SELF

send_email = True
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
IF MX_region = "TRAINING" THEN send_email = False    'Will default to sending email, but not if region is in training region.

Call MAXIS_background_check

'PRIV handling
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv) 'navigating to stat memb to gather the ref number and name.
IF is_this_priv = TRUE THEN script_end_procedure("PRIV case, cannot access/update. The script will now end.")


'Selecting the HH member
DO
    CALL custom_HH_member_custom_dialog(HH_member_array)
    IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

'Navigate back to STAT/MEMB in case user has navigated away
Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "MEMB")

'--------------------------------------------------------------------------------Gathering the MEMB/ALIA information
'Establishing array
uc_membs = 0       'incrementor for array
DIM uc_members_array()  'Declaring the array this is what this list is
ReDim uc_members_array(client_alia_ssn_const, 0)  'Resizing the array 'redimmed to the size of the last constant  'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const maxis_case_number_const  	 	= 0 '=  Maxis'
const member_number_const   		= 1 '=  Member Number
const client_first_name_const       = 2 '=  First Name MEMB
const client_last_name_const        = 3 '=  Last Name MEMB
const client_mid_name_const    	    = 4 '=  Middle initial MEMB
const client_DOB_const   		    = 5 '=  Date of Birth MEMB
const client_ssn_const		        = 6 '=  SSN
const client_age_const	            = 7 '=  age MEMB
const client_alia_name_const	    = 8 '=  ALIA Name
const client_alia_ssn_const		    = 9 '=  ALIA SSN

FOR EACH person IN HH_member_array
	CALL write_value_and_transmit(person, 20, 76) 'reads the reference number, last name, first name, and THEN puts it into an array YOU HAVENT defined the uc_members_array yet

	script_run_lowdown = script_run_lowdown & vbCr & "STAT-MEMB  --  person: " & person
	for memb_row = 1 to 24
		EMReadScreen mx_line, 79, memb_row, 1
		script_run_lowdown = script_run_lowdown & vbCr & mx_line
	next
	script_run_lowdown = script_run_lowdown & vbCr & "=============================================================" & vbCr & vbCr

	EMReadscreen ref_nbr, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EMReadscreen client_SSN, 11, 7, 42
    If client_ssn = "___ __ ____" THEN client_ssn = ""

	last_name = trim(replace(last_name, "_", "")) & " "
    first_name = trim(replace(first_name, "_", "")) & " "
    mid_initial = replace(mid_initial, "_", "")
    EMReadScreen client_age, 2, 8, 76
    IF client_age = "  " THEN client_age = 0
    client_age = client_age * 1
    ReDim Preserve uc_members_array(client_alia_ssn_const, uc_membs)  'redimmed to the size of the last constant
    uc_members_array(member_number_const,     uc_membs) = ref_nbr
    uc_members_array(client_first_name_const, uc_membs) = first_name
    uc_members_array(client_last_name_const,  uc_membs) = last_name
    uc_members_array(client_mid_name_const,   uc_membs) = mid_initial
    uc_members_array(client_DOB_const,        uc_membs) = client_DOB
    uc_members_array(client_ssn_const,        uc_membs) = client_SSN
    uc_members_array(client_age_const,        uc_membs) = client_age
    uc_membs = uc_membs + 1
    STATS_counter = STATS_counter + 1
NEXT
'--------------------------------------------------------------------------------ALIA
CALL navigate_to_MAXIS_screen("STAT", "ALIA")

FOR uc_membs = 0 to uBound(uc_members_array, 2)
	CALL write_value_and_transmit(uc_members_array(member_number_const, uc_membs), 20, 76)
    row = 7
    DO
        EMReadscreen alia_ref_num, 02, 04, 33
	    EMReadScreen alia_last_name, 17, row, 26
        alia_last_name = replace(alia_last_name, "_", "")
        alia_last_name = trim(alia_last_name)
        EMReadScreen alia_first_name, 12, row, 53
        alia_first_name = replace(alia_first_name, "_", "")
        alia_first_name = trim(alia_first_name)
        EMReadScreen mid_initial, 1, row, 75
        IF alia_last_name = "" THEN EXIT DO
        row = row + 1
        uc_members_array(client_alia_name_const, uc_membs) = uc_members_array(client_alia_name_const, uc_membs) & ", " & alia_first_name & " " & alia_last_name
    Loop until row = 13

    row = 15
    Do
        EMReadScreen alia_client_SSN, 11, row, 28
        If alia_client_SSN = "___ __ ____" then
            alia_client_SSN = ""
            EXIT DO
        ELSE
            alia_client_SSN = trim(alia_client_SSN)
        END IF
        uc_members_array(client_alia_ssn_const, uc_membs) = uc_members_array(client_alia_ssn_const, uc_membs) & ", " & alia_client_SSN 'adding the varibale without writing over it each time i go throught the loop'
        EMReadScreen alia_client_SSN_II, 11, row, 53
        If alia_client_SSN_II = "___ __ ____" then
            alia_client_SSN_II = ""
            EXIT DO
        END IF
        row = row + 1
        uc_members_array(client_alia_ssn_const, uc_membs) = uc_members_array(client_alia_ssn_const, uc_membs) & ", " & alia_client_SSN_II
    Loop until row = 18

    IF left(uc_members_array(client_alia_name_const, uc_membs), 1) = "," THEN uc_members_array(client_alia_name_const, uc_membs) = right(uc_members_array(client_alia_name_const, uc_membs), len(uc_members_array(client_alia_name_const, uc_membs)) - 1)
    uc_members_array(client_alia_name_const, uc_membs) = trim(uc_members_array(client_alia_name_const, uc_membs)) 'once I have added everything to the array THEN i can format'
    IF left(uc_members_array(client_alia_ssn_const, uc_membs), 1) = "," THEN uc_members_array(client_alia_ssn_const, uc_membs) = right(uc_members_array(client_alia_ssn_const, uc_membs), len(uc_members_array(client_alia_ssn_const, uc_membs)) - 1)
    uc_members_array(client_alia_ssn_const, uc_membs) = trim(uc_members_array(client_alia_ssn_const, uc_membs))
    alia_first_name = ""
    alia_last_name = ""
    alia_client_SSN = ""
    alia_client_SSN_II = ""
NEXT

email_bzst = False
FOR uc_membs = 0 to Ubound(uc_members_array, 2) 'start at the zero person and go to each of the selected people '
	If mid(uc_members_array(client_ssn_const,  uc_membs), 5, 1) = " " Then email_bzst = True
    member_info = member_info & "-------------"  & vbNewLine  & "Name of Resident: " & uc_members_array(client_first_name_const, uc_membs) & " " & uc_members_array(client_mid_name_const, uc_membs) & " " & uc_members_array(client_last_name_const, uc_membs)  &  vbCr & "DOB: " & uc_members_array(client_DOB_const,  uc_membs) &  vbcr & "SSN of Resident: " & uc_members_array(client_ssn_const,  uc_membs)
    IF trim(uc_members_array(client_alia_name_const, uc_membs)) <> "" THEN member_info = member_info & vbNewLine &  "ALIA Name: " & uc_members_array(client_alia_name_const, uc_membs)
    If trim(uc_members_array(client_alia_ssn_const,  uc_membs)) <> "" THEN member_info = member_info & vbNewLine & "ALIA SSN: " & uc_members_array(client_alia_ssn_const,  uc_membs)
    member_info = member_info & vbNewLine & "Please review deductions and withholdings for this individual. " & vbNewLine
NEXT

IF cca_checkbox = CHECKED THEN member_info = "CCA Request" & vbNewLine & member_info
IF other_checkbox = CHECKED and other_check_editbox <> "" THEN member_info = "Other Request: " & other_check_editbox & vbNewLine & member_info

CALL find_user_name(the_person_running_the_script)' this is for the signature in the email'

If email_bzst = True Then
	bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
	subject_of_email = "UC Verif Numbers reversed?! Case " & MAXIS_case_number & " (Automated Report)"
	full_text = "Case Number: " & MAXIS_case_number
	full_text = full_text & vbCr & "Script Run occurred on " & date & " at " & time & vbCr

	full_text = full_text & vbCr & "Member INFO:" & vbCr & member_info
	full_text = full_text & vbCr & vbCr & "Script Run Lowdown:" & vbCr & script_run_lowdown
	full_text = full_text & vbCr & "Submitted By: " & the_person_running_the_script
	attachment_here = ""
	Call create_outlook_email("", bzt_email, "", "", subject_of_email, 1, False, "", "", False, "", full_text, True, attachment_here, True)
End If

'Creates message box with email information if using the training region.
email_information = "Email Information:" & vbcr & vbcr & "UC Request for Case #" & MAXIS_case_number & _
vbcr & vbcr & "Member Info: " & member_info & _
vbcr & vbcr & "Submitted By: " & the_person_running_the_script

'Call create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
IF send_email = TRUE THEN Call create_outlook_email("", "HSPH.ES.DEED@Hennepin.us", "", "", "UC Request for Case #" & MAXIS_case_number, 1, False, "", "", False, "", member_info & vbNewLine & vbNewLine & "Submitted By: " & vbNewLine & the_person_running_the_script, False, "", True)
'will create email, will send.

IF MX_region = "TRAINING" then msgbox email_information

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/19/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/19/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/19/2022
'--All variables in dialog match mandatory fields-------------------------------05/19/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/19/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/19/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/19/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/19/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------05/19/2022
'--PRIV Case handling reviewed -------------------------------------------------05/19/2022
'--Out-of-County handling reviewed----------------------------------------------05/19/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/19/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/19/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/19/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/19/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/19/2022
'--Denomination reviewed -------------------------------------------------------05/19/2022
'--Script name reviewed---------------------------------------------------------05/19/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/19/2022-------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/19/2022
'--comment Code-----------------------------------------------------------------05/19/2022
'--Update Changelog for release/update------------------------------------------05/19/2022
'--Remove testing message boxes-------------------------------------------------05/19/2022
'--Remove testing code/unnecessary code-----------------------------------------05/19/2022
'--Review/update SharePoint instructions----------------------------------------05/19/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/19/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/19/2022
'--Complete misc. documentation (if applicable)---------------------------------05/19/2022
'--Update project team/issue contact (if applicable)----------------------------05/19/2022
