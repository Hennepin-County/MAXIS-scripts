'GATHERING STATS===========================================================================================
name_of_script = "NOTES - AVS.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 240
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
call changelog_update("03/27/2024", "Updated the name of AVS to Asset Verification Service.", "Dave Courtright, Hennepin County")
call changelog_update("02/22/2024", "Enabled the Renewal option for submitting an AVS Request.", "Ilse Ferris, Hennepin County")
call changelog_update("06/26/2023", "Disabled renewal option due to asset diregard through 05/31/2024.", "Ilse Ferris, Hennepin County")
call changelog_update("05/01/2023", "Updated AVS Portal Case Review reminder for 11th day after submitting an Ad-Hoc request.", "Mark Riegel, Hennepin County")
call changelog_update("03/20/2023", "Added Change in Basis option, and enabled the Renewal option for submitting an AVS Request.", "Ilse Ferris, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("12/30/2022", "Fixed inhibiting bug if HH members do not have an age listed on STAT/MEMB.", "Ilse Ferris, Hennepin County")
call changelog_update("05/10/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
closing_msg = "Success! Your AVS case note has been created. Please review for accuracy & any additional information."  'initial closing message. This may increment based on options selected.
get_county_code
Call Check_for_MAXIS(False)

'Adding case number if using the script from the DAIL scrubber
If MAXIS_case_number = "" then
    EmReadscreen DAIL_panel, 4, 2, 48
    If DAIL_panel = "DAIL" then
        EmReadscreen MAXIS_case_number, 8, 5, 73
        MAXIS_case_number = trim(MAXIS_case_number)
        'Defaulting initial option based on the dail message.
        EMReadScreen full_message, 60, 6, 20
        full_message = trim(full_message)
        If Instr(full_message, "AVS 10-DAY CHECK IS DUE") then
            initial_option = "AVS Submission/Results"
        else
            initial_option = "AVS Forms"
        End if
    End if
End if

'----------------------------------------------------------------------------------------------------Initial dialog
initial_help_text = "*** What is the AVS? ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & _
"The Asset Verification Service (AVS) is a web-based service that provides information and verification of some accounts held in financial institutions. It does not provide information on property assets such as cars or homes. AVS must be used once at application, and when a person changes to a Medical Assistance for People Who Are Age 65 or Older and People Who Are Blind or Have a Disability (MA-ABD) basis of eligibility and are subject to an asset test. It is also used at renewal to verify certain assets." & vbNewLine & vbNewLine & _
"If a resident is applying for an MHCP without an asset test then the AVS should not be run. This verification is not meant for any other public assistance programs besides health care."

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 85, "AVS Initial Selection Dialog"
  EditBox 75, 10, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
  PushButton 135, 10, 10, 15, "!", initial_help_button
  DropListBox 75, 30, 105, 15, "Select one..."+chr(9)+"AVS Forms"+chr(9)+"AVS Submission/Results", initial_option
  DropListBox 75, 45, 105, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change In Basis"+chr(9)+"Renewal", HC_process
  ButtonGroup ButtonPressed
    OkButton 75, 65, 40, 15
    CancelButton 120, 65, 40, 15
  Text 5, 35, 70, 10, "Select AVS Process:"
  Text 10, 50, 65, 10, "Select HC Process:"
  Text 25, 15, 45, 10, "Case number:"
EndDialog

'Initial dialog: user will input case number and initial options
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1
		cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
        If ButtonPressed = initial_help_button then
            tips_tricks_msg = MsgBox(initial_help_text, vbInformation, "Tips and Tricks") 'see initial_help_text above for details of the text
            err_msg = "LOOP" & err_msg
        End if
    	Call validate_MAXIS_case_number(err_msg, "*")
        If initial_option = "Select one..." then err_msg = err_msg & vbcr & "* Select the AVS process."
        If HC_process = "Select one..." then err_msg = err_msg & vbcr & "* Select the health care process."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

EmReadscreen county_code, 4, 21, 21
If county_code <> UCASE(worker_county_code) then script_end_procedure("This case is an out-of-county case, and cannot be case noted. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member/AREP/Sponsor information for signature selection array
'Setting up main array
avs_membs = 0       'incrementor for array
Dim avs_members_array()
ReDim avs_members_array(memb_array_last_const, 0)   'redimmed to the size of the last constant
const member_number_const      = 0
const member_info_const        = 1
const member_name_const        = 2
const marital_status_const     = 3
const checked_const            = 4
const hc_applicant_const       = 5
const applicant_type_const     = 6
const forms_status_const       = 7
const avs_status_const         = 8
const request_type_const       = 9
const avs_results_const        = 10
const avs_returned_notes_const = 11
const avs_date_const           = 12
const accounts_verified_const  = 13
const unreported_assets_const  = 14
const ECF_const                = 15
const additional_info_const    = 16
const status_msg_const         = 17
const form_type_const          = 18 
const auth_date_const          = 19
const form_valid_const         = 20
const auth_sent_date_const     = 21
const sigs_needed_const        = 22
const avs_action_const         = 23 
const first_submitted_const    = 24
const second_submitted_const   = 25
const memb_smi_const           = 26
const ad_hoc_type_const        = 27 
const ad_hoc_sent_count_const  = 28
const ad_hoc_sent_date_const   = 29
const ad_hoc_reviewed_date_const = 30 
const ad_hoc_closed_date_const = 31
const ad_hoc_status_const      = 32
const memb_age_const           = 33
const memb_array_last_const    = 34
add_to_array = False    'defaulting to false
DO
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EmReadscreen relationship_code, 2, 10, 42
    EmReadscreen client_age, 3, 8, 76
    EmReadscreen client_ssn, 11, 7, 42
    EmReadScreen client_smi, 9, 5, 46
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
    client_age = trim(client_age)

    If relationship_code = "01" then add_to_array = True    'applicants of HC yes
    If relationship_code = "02" then add_to_array = True    'spouses of HC applicants yes

    If client_ssn = "___ __ ____" then
        client_ssn = ""
        'Folks who have no ssn are not required to submit an AVS inquiry
    Else
        client_ssn = replace(client_ssn, " ", "")
    End if

    If client_age = "" then
        add_to_array = False    'Added if statement to resolve inhibiting script issues
    ElseIf client_age < 21 then
        add_to_array = False  'under 21 are not required to sign per EPM 2.3.3.2.1 Asset Limits
    End if

    If add_to_array = True then
        ReDim Preserve avs_members_array(memb_array_last_const, avs_membs)  'redimmed to the size of the last constant
        avs_members_array(member_info_const,    avs_membs) = ref_nbr & " " & last_name & first_name
        avs_members_array(member_number_const,  avs_membs) = ref_nbr
        avs_members_array(member_name_const,    avs_membs) = first_name & "" & last_name
        avs_members_array(memb_smi_const,       avs_membs) = client_smi
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
		avs_members_array(memb_age_const,       avs_membs) = client_age
        If client_ssn = "" then
			If initial_option = "AVS Submission/Results" then avs_members_array(request_type_const, avs_membs) = "N/A - No SSN" 'NO SSN need to sign forms, but we just need to case note AVS Submission exemption
		End if
        avs_membs = avs_membs + 1
    End if
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'Handling in case no members are identified as needing the form. Helping to reduce errors for workers.
If avs_membs = 0 then script_end_procedure_with_error_report("No members on this case are required by policy to sign the AVS form. Please review case if necessary.")

Call navigate_to_MAXIS_screen("STAT", "MEMI")   'Finding marital status to add to array
For items= 0 to Ubound(avs_members_array, 2)
    EmWriteScreen avs_members_array(member_number_const, item), 20, 76
    transmit
    EmReadscreen marital_status, 1, 7, 40
    avs_members_array(marital_status_const, item) = marital_status
Next

'----------------------------------------------------------------------------------------------------SPONSOR information if applicable to be added to array
Call navigate_to_MAXIS_screen("STAT", "SPON")
EmReadscreen total_spon_panels, 1, 2, 78
Do
    If total_spon_panels = "0" then exit do
    EMReadScreen spon_name, 20, 8, 38
    spon_name = replace(spon_name, "_", "")
    If trim(spon_name) <> "" then
        'Adding to main array
        ReDim Preserve avs_members_array(additional_info_const,     avs_membs)
        avs_members_array(member_info_const,    avs_membs) = "Sponsor - " & trim(spon_name)
        avs_members_array(member_name_const,    avs_membs) = trim(spon_name)
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_members_array(hc_applicant_const,   avs_membs) = FALSE      'AREP's will not be applicants
        avs_members_array(applicant_type_const, avs_membs) = "Deeming"
        avs_membs = avs_membs + 1
    End if
    transmit
    EMReadScreen last_panel, 5, 24, 2
Loop until last_panel = "ENTER"	'This means that there are no other faci panels

'STAT/TYPE info is to help default some of the initial options for the applicant type.
Call navigate_to_MAXIS_screen("STAT", "TYPE")
For items= 0 to Ubound(avs_members_array, 2)
    If avs_members_array(hc_applicant_const, item) = "" then
        'adding TYPE Information to be output into the dialog and case note
        row = 6
        Do
            EmReadscreen type_memb_number, 2, row, 3
            EmReadscreen applicant_type, 1, row, 37
            If type_memb_number = avs_members_array(member_number_const, item) then
                If applicant_type = "Y" then
                    avs_members_array(hc_applicant_const, item) = True
                    If avs_members_array(marital_status_const, item) = "M" then
                        avs_members_array(applicant_type_const, item) = "Applying/Spouse"
                    Else
                        avs_members_array(applicant_type_const, item) = "Applying"
                    End if
                    exit do
                Elseif applicant_type = "N" OR applicant_type  = "_" then
                    avs_members_array(hc_applicant_const, item) = False
                    If avs_members_array(marital_status_const, item) = "M" then
                        avs_members_array(applicant_type_const, item) = "Spouse"
                    Else
                        avs_members_array(applicant_type_const, item) = "Not Applying"
                    End if
                    exit do
                End if
            Else
                row = row + 1
            End if
        Loop until trim(type_memb_number) = ""
	End if
Next

Do
    'Blanking out variables for the array to start the AVS Submission process. This is a valid option for users to select is the AVS Forms process is selected, and the form(s) are returned complete.
    If confirm_msgbox = vbYes then
        For items= 0 to ubound(avs_members_array, 2)
            run_initial_option = False
            avs_members_array(forms_status_const,       item) = ""
            avs_members_array(avs_status_const,         item) = ""
            avs_members_array(request_type_const,       item) = ""
            avs_members_array(avs_results_const,        item) = ""
            avs_members_array(avs_returned_notes_const, item) = ""
            avs_members_array(avs_date_const,           item) = ""
            avs_members_array(accounts_verified_const,  item) = ""
            avs_members_array(unreported_assets_const,  item) = ""
            avs_members_array(ECF_const,                item) = ""
            avs_members_array(additional_info_const,    item) = ""
        Next
    End if
    '----------------------------------------------------------------------------------------------------SELECTING AVS MEMBERS: Based on who is required to sign form/submit AVS
    'Text for the next dialogs based on initial option selected by the user
    If initial_option = "AVS Forms" then
        selection_text = "Select all members REQUIRED to sign AVS form(s):"
        type_text = "Applicant"
        dialog_text = "Forms"
        help_button_text = "*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbcr & vbcr & _
        "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _
        "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _
        "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country." & vbNewLine & vbNewLine & _
        "Information Source: DHS-7823 Form - Authorization to Obtain Financial Information from the Asset Verification Service (AVS)."

        help_button_2_text = "*** What date should I enter here? ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & _
        "The Form Status will determine what you will enter in this field." & vbNewLine & vbNewLine & _
        "- Initial Request: Enter the date the request was sent to the resident." & vbNewLine & _
        "- Not Received: Enter the status date (most likely today's date)." & vbNewLine & _
        "- Received - Complete or Received - Incomplete: Enter the date received in the agency."
    End if

    If initial_option = "AVS Submission/Results" then
        selection_text = "Select all members REQUIRED to sign AVS form(s):"
        type_text = "Request"
        dialog_text = "AVS"

        help_button_text = "*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbcr & vbcr & _
        "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _
        "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _
        "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country." & vbNewLine & vbNewLine & _
        "Information Source: DHS-7823 Form - Authorization to Obtain Financial Information from the Asset Verification Service (AVS)."

        help_button_2_text = "*** What date should I enter here? ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & _
        "The AVS Status will determine what you will enter in this field. These will usually be the current date." & vbNewLine & vbNewLine & _
        "- Submitting a Request: Enter the date the request was sent in the AVS system." & vbNewLine & _
        "- Review Results or Results After Decision: Enter the date the results were reviewed in the AVS system."
    End if

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 200, (50 + (items* 20)), "AVS Member Selection Dialog"
        Text 5, 5, 180, 10, selection_text
        ButtonGroup ButtonPressed
        PushButton 170, 0, 10, 15, "!", help_button
        For items= 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 1 then checkbox 15, (20 + (items* 20)), 130, 15, avs_members_array(member_info_const, item), avs_members_array(checked_const, item)
        Next
        ButtonGroup ButtonPressed
        OkButton 85, (30 + (items * 20)), 45, 15
        CancelButton 135, (30 + (items * 20)), 45, 15
    EndDialog

    'Member selection Dialog
    Do
        Do
            err_msg = ""
            Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
            cancel_without_confirmation
            If ButtonPressed = help_button then
                tips_tricks_msg = MsgBox(help_button_text, vbInformation, "Tips and Tricks") 'see help_button_text above for details of the text
                err_msg = "LOOP" & err_msg
            End if
            'ensuring that users have
            checked_count = 0
            FOR items = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
                If avs_members_array(checked_const, item) = 1 then checked_count = checked_count + 1 'Ignores and blank scanned in persons/strings to avoid a blank checkbox
            NEXT
            If checked_count = 0 then err_msg = err_msg & vbcr & "* Select all persons responsible for signing the AVS form."
            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    'Resizing the array based on who was selected in the previous dialog. Revaluing the array if selected or checked in the previous dialog.
    resize_counter = 0
    For items = 0 to UBound(avs_members_array, 2)
        If avs_members_array(checked_const, item) = 1 Then
            avs_members_array(member_number_const     , resize_counter) = avs_members_array(member_number_const     , item)
            avs_members_array(member_info_const       , resize_counter) = avs_members_array(member_info_const       , item)
            avs_members_array(member_name_const       , resize_counter) = avs_members_array(member_name_const       , item)
            avs_members_array(marital_status_const    , resize_counter) = avs_members_array(marital_status_const    , item)
            avs_members_array(checked_const           , resize_counter) = avs_members_array(checked_const           , item)
            avs_members_array(hc_applicant_const      , resize_counter) = avs_members_array(hc_applicant_const      , item)
            avs_members_array(applicant_type_const    , resize_counter) = avs_members_array(applicant_type_const    , item)
            avs_members_array(forms_status_const      , resize_counter) = avs_members_array(forms_status_const      , item)
            avs_members_array(avs_status_const        , resize_counter) = avs_members_array(avs_status_const        , item)
            avs_members_array(request_type_const      , resize_counter) = avs_members_array(request_type_const      , item)
            avs_members_array(avs_results_const       , resize_counter) = avs_members_array(avs_results_const       , item)
            avs_members_array(avs_returned_notes_const, resize_counter) = avs_members_array(avs_returned_notes_const, item)
            avs_members_array(avs_date_const          , resize_counter) = avs_members_array(avs_date_const          , item)
            avs_members_array(accounts_verified_const , resize_counter) = avs_members_array(accounts_verified_const , item)
            avs_members_array(unreported_assets_const , resize_counter) = avs_members_array(unreported_assets_const , item)
            avs_members_array(ECF_const               , resize_counter) = avs_members_array(ECF_const               , item)
            avs_members_array(additional_info_const   , resize_counter) = avs_members_array(additional_info_const   , item)
            avs_members_array(memb_smi_const          , resize_counter) = avs_members_array(memb_smi_const          , item)
            avs_members_array(memb_age_const          , resize_counter) = avs_members_array(memb_age_const          , item)
            resize_counter = resize_counter + 1
            STATS_counter = STATS_counter + 1
        End If
    Next
    resize_counter = resize_counter - 1
    ReDim Preserve avs_members_array(memb_array_last_const, resize_counter) 'rediming the array to move forward with the selected members.
    call generate_client_list(member_list, "")

    'Connecting to the database and finding current info on this person
    SQL_Case_Number = right("00000000" & MAXIS_case_number, 8)
	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
    For this_memb = 0 to UBound(avs_members_array, 2) 'Check data table for existing records for each member	
        objSQL = "SELECT * FROM ES.ES_AVSList WHERE CaseNumber = '" & SQL_Case_Number & "'	AND SMI = '" & avs_members_array(memb_smi_const, this_memb) & "'"	'Find the record matching case / SMI
        
        objRecordSet.Open objSQL, objConnection
            If Not objRecordSet.bof Then 'If we have an existing record for this member, read the values
                avs_members_array(auth_date_const, this_memb)              = objRecordSet("AVSFormDate")
                avs_members_array(form_type_const , this_memb)  		   = objRecordSet("AVSFormType")
                avs_members_array(form_valid_const, this_memb)             = objRecordSet("AVSFormValid")
                avs_members_array(auth_sent_date_const, this_memb)         = objRecordSet("AuthSentDate")
                avs_members_array(ad_hoc_type_const, this_memb)            = objRecordSet("AdHocType")
                avs_members_array(ad_hoc_sent_count_const, this_memb)      = objRecordSet("AdHocSentCount")
                avs_members_array(ad_hoc_sent_date_const, this_memb)       = objRecordSet("AdHocSentDate")
                avs_members_array(ad_hoc_reviewed_date_const, this_memb)   = objRecordSet("AdHocReviewedWorker") 
                avs_members_array(ad_hoc_closed_date_const, this_memb)     = objRecordSet("AdHocClosedDate")
            End If 
        ObjRecordSet.Close
    Next 
	objConnection.close 'close down the connection
   
    'Set the current status / info message for each member
    For this_memb = 0 to UBound(avs_members_array, 2)
        If isdate(avs_members_array(ad_hoc_sent_date_const, this_memb)) = True Then '                    
            If datediff("d", avs_members_array(ad_hoc_sent_date_const, this_memb), date) > 90 Then 
                avs_members_array(ad_hoc_status_const, this_memb) = "Last ad-hoc recorded was " & avs_members_array(ad_hoc_sent_date_const, this_memb) 
            ElseIf avs_members_array(ad_hoc_reviewed_date_const, this_memb) = "" AND avs_members_array(ad_hoc_sent_date_const, this_memb) <> "" Then
                 msgbox "also...."
                 avs_members_array(status_msg_const, this_memb) = "Ad hoc request sent " & avs_members_array(ad_hoc_sent_date_const, this_memb) & ". Review ad-hoc results."
                 avs_members_array(ad_hoc_status_const, this_memb) = "Sent " & avs_members_array(ad_hoc_sent_date_const, this_memb)
            ElseIf avs_members_array(ad_hoc_reviewed_date_const, this_memb) <> "" AND avs_members_array(ad_hoc_closed_date_const, this_memb) = "" Then
                avs_members_array(ad_hoc_status_const, this_memb) = "Reviewed " & avs_members_array(ad_hoc_reviewed_date_const, this_memb)
            Else
                avs_members_array(status_msg_const, this_memb) = "Check ECF Case file for valid authorization for member, update below or send forms if needed."
            End If 
        Else 'No sent date for Ad hoc
            avs_members_array(ad_hoc_status_const, this_memb) = "Process not started."
            If avs_members_array(form_valid_const, this_memb) = "1" Then 
                avs_members_array(status_msg_const, this_memb) = "Valid forms on file. Send AVS ad hoc request(s) for this member."  
            Else 
                avs_members_array(status_msg_const, this_memb) = "Check ECF Case file for valid authorization for member, update below or send forms if needed."
            End If 
        End If 
    Next 
    

    '----------------------------------------------------------------------------------------------------Adding in information about the AVS Members selected
    If HC_process = "Renewal" Then
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 575, (55 + (checked_count * 65)), "AVS At Renewal"
        y_pos = 10

         For this_memb = 0 to UBound(avs_members_array, 2)	
          Text 45, y_pos + 10, 340, 10, avs_members_array(status_msg_const, this_memb) '"Check case file for existing authorizations and complete below with status of any authorization on file."
            'TODO 
         GroupBox 10, y_pos, 560, 65, avs_members_array(member_number_const, this_memb) & " " & avs_members_array(member_name_const, this_memb)
         Text 15, y_pos+30, 55, 10, "AVS Form Type: "
         DropListBox 70, y_pos+25, 75, 15, ""+chr(9)+"DHS-7823 Auth Form"+chr(9)+"HCAPP"+chr(9)+"HC Renewal", avs_members_array(form_type_const, this_memb)
         EditBox 190, y_pos+25, 45, 15, avs_members_array(auth_date_const, this_memb)
         Text 150, y_pos+30, 40, 10, "Form Date:"
         DropListBox 265, y_pos+25, 75, 15, ""+chr(9)+"Valid form on file"+chr(9)+"Form is invalid"+chr(9)+"No form on file", avs_members_array(form_valid_const, this_memb)
         Text 240, y_pos+30, 25, 10, "Status:"
         Text 350, y_pos+25, 70, 20, "Auth form sent,     must be signed by:"
         EditBox 415, y_pos+25, 150, 15, avs_members_array(sigs_needed_const, this_memb)

         Text 15, y_pos+50, 50, 10, "Ad Hoc status:"
         DropListBox 190, y_pos+45, 50, 10, ""+chr(9)+"Submitted"+chr(9)+"Reviewed "+chr(9)+"Closed", avs_members_array(avs_action_const, this_memb)
         Text 70, y_pos+50, 80, 10, avs_members_array(ad_hoc_status_const, this_memb)
         Text 155, y_pos+50, 30, 10, "Action:"
         Text 245, y_pos+50, 15, 10, "For:"
         DropListBox 265, y_pos+45, 125, 15, member_list, avs_members_array(first_submitted_const, this_memb)
         DropListBox 400, y_pos+45, 115, 15, member_list, avs_members_array(second_submitted_const, this_memb)
         ButtonGroup ButtonPressed
           PushButton 520, y_pos+45, 45, 10, "Sponsors", Button5
        y_pos = y_pos + 70
        Next
         y_pos = y_pos + 10
         Text 15, y_pos+5, 45, 15, "Other Notes:"
         EditBox 60, y_pos, 225, 15, other_notes
         Text 290, y_pos+5, 60, 15, "Worker Signature:"
         EditBox 350, y_pos, 120, 15, worker_signature
        
           OkButton 475, y_pos, 40, 15
           CancelButton 515, y_pos, 40, 15
        EndDialog

        Do
            Do
                err_msg = ""
                Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
                cancel_confirmation
                If ButtonPressed = help_button_1 then
                    tips_tricks_msg = MsgBox(help_button_text, vbInformation, "Tips and Tricks")
                    err_msg = "LOOP" & err_msg
                End if
                If ButtonPressed = help_button_2 then
                    tips_tricks_msg = MsgBox(help_button_2_text, vbInformation, "Tips and Tricks")
                    err_msg = "LOOP" & err_msg
                End if

                'mandatory fields for all AVS_membs
                FOR this_thing = 0 to UBound(avs_members_array, 2)
                    'AVS Forms mandatory fields

                    'AVS Submission/Results mandatory fields

                    'If trim(avs_members_array(avs_date_const, item)) = "" or isdate(avs_members_array(avs_date_const, item)) = False then err_msg = err_msg & vbcr & "* Enter the " & dialog_text & " status date for: " & avs_members_array(member_info_const, item)
                NEXT
                If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Enter your worker signature."
                IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
            LOOP UNTIL err_msg = ""
            Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in
        
        'This section updates SQL table with status from dialog
        Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	    Set objRecordSet = CreateObject("ADODB.Recordset")

	    'opening the connections and data table
	    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
        For this_memb = 0 to UBound(avs_members_array, 2) 'Loop through each member and update data table
            'set variables
            form_type = avs_members_array(form_type_const, this_memb)
            memb_smi  = avs_members_array(memb_smi_const, this_memb)
                       
            objSQL = "SELECT * FROM ES.ES_AVSList WHERE CaseNumber = '" & SQL_Case_Number & "'	AND SMI = '" & memb_smi & "'"	'Find the record matching case / SMI
            objRecordSet.Open objSQL, objConnection
            
            'check if data exists before update
            If objRecordSet.Bof and objRecordSet.EOF Then 'There isn't an existing record for this person, we will insert one
               msgbox "Record not found"
                If avs_members_array(memb_age_const, this_memb)*1 >= 21 Then
					memb_asset_test = 1
				Else
					memb_asset_test = 0
				End If 
				
				this_month = datepart("M", date)
				If len(this_month) = 1 Then this_month = "0" & this_month
				year_month = datepart("YYYY", date) & this_month
                
                objAVSinsert =  "INSERT INTO ES.ES_AVSList (YearMonth, SMI, CaseNumber, AssetTest) VALUES ('" & year_month & "', '" & memb_smi & "', '" & SQL_Case_Number & "', '" & memb_asset_test & "')"
                
			    Set objinsertRecordSet = CreateObject("ADODB.Recordset")
    			'Opening and inserting the values
				objinsertRecordSet.Open objAVSinsert, objUpdateConnection	
            End If  
            
            'AVS Form updates

        msgbox "what?"

            'AVS sent updates
            If objRecordSet("AdHocSentDate") <> avs_members_array(ad_hoc_sent_date_const, this_memb) Then
                ObjSentUpdate = "UPDATE ES.ES_AVSList SET AdHocSentDate = '" & avs_members_array(ad_hoc_sent_date_const, this_memb)  & "', "&_     
                "AdhocSentCount  = '" & avs_members_array(ad_hoc_sent_count_const, this_memb) & "', "&_
                "AdHocSentWorker = '" & user_ID_for_validation & "', "&_
                "WHERE CaseNumber = '" & SQL_Case_Number & "' and SMI = '" & memb_smi & "'"
			    objRecordSet.Open objSentUpdate, objUpdateConnection           
            End If 
             'AVS closed updates
            If objRecordSet("AdHocClosedDate") <> avs_members_array(ad_hoc_closed_date_const, this_memb) Then
                ObjClosedUpdate = "UPDATE ES.ES_AVSList SET AdHocClosedDate = '" & avs_members_array(ad_hoc_closed_date_const, this_memb)  & "', "&_     
                " AdHocClosedWorker = '" & user_ID_for_validation & "', "&_
                "WHERE CaseNumber = '" & SQL_Case_Number & "' and SMI = '" & memb_smi & "'"
			    objRecordSet.Open objClosedUpdate, objUpdateConnection           
            End If 
                   '    AdHocType =  '" & avs_members_array(ad_hoc_type_const, this_memb)  & "', &_
                
                'objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"
            'AVS form updates
            If avs_members_array(form_valid_const, this_memb) = "Valid form on file" Then 
                form_valid = "1"
            Else
                form_valid = "0"
            End If 

            If objRecordSet("AVSFormValid") <> form_valid OR  objRecordSet("AVSFormDate") <> avs_members_array(auth_date_const, this_memb) OR objRecordSet("AVSFormType") <> avs_members_array(form_type_const , this_memb) Then 
                msgbox objRecordSet("AVSFormValid") 
                objFormUpdateSQL = "UPDATE ES.ES_AVSList SET AVSFormDate = '" & avs_members_array(auth_date_const, this_memb) & "', "&_
                "AVSFormType = '" & avs_members_array(form_type_const , this_memb) & "', "&_                
                "AVSFormValid = '" &  form_valid & "', "&_
                "WHERE CaseNumber = '" & SQL_Case_Number & "' and SMI = '" & memb_smi & "'"
                objRecordSet.Open objFormUpdateSQL, objUpdateConnection 
            End If 
            'AVS
            If objRecordSet("AuthSentDate") <> avs_members_array(auth_sent_date_const, this_memb) Then
                objAuthSQL = "UPDATE ES.ES_AVSList SET AuthSentDate = '" &   avs_members_array(auth_sent_date_const, this_memb) & "', "&_           
                "WHERE CaseNumber = '" & SQL_Case_Number & "' and SMI = '" & memb_smi & "'"
                objRecordSet.Open objAuthSQL, objUpdateConnection
            End If 
        Next 
	    objConnection.close 'close down the connection
        msgbox "Stop now dude"
    Else 
         Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 575, (115 + (checked_count * 15)), "AVS Member Information Dialog"
        GroupBox 10, 5, 550, (60 + (checked_count * 15)), "Complete the following information for required AVS members:"
        Text 20, 25, 520, 10, "----------AVS Member-------------------------------------" & type_text & " Type----------------------" & dialog_text & " Status-------------------" & dialog_text & " Sent/Rec'd Date-------------------Person-Based Info----------------"
          For items = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
              y_pos = (50 + items * 20)
              Text 20, y_pos, 130, 15, avs_members_array(member_info_const, items)
              'AVS Forms selections
              If initial_option = "AVS Forms" then
                  DropListBox 150, y_pos - 5, 70, 15, "Select one..."+chr(9)+"Applying"+chr(9)+"Applying/Spouse"+chr(9)+"Deeming"+chr(9)+"Not Applying"+chr(9)+"Spouse", avs_members_array(applicant_type_const, items)
                  DropListBox 225, y_pos - 5, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", avs_members_array(forms_status_const, item)
              End if
              'AVS Submission/Results selections
              If initial_option = "AVS Submission/Results" then
                      DropListBox 135, y_pos - 5, 90, 15, "Select one..."+chr(9)+"BI - Brain Injury Waiver"+chr(9)+"BX - Blind"+chr(9)+"CA - CAC Waiver"+chr(9)+"CD - CADI Waiver"+chr(9)+"DD - DD Waiver"+chr(9)+"DP - MA-EPD"+chr(9)+"DX - Disability"+chr(9)+"EH - EMA"+chr(9)+"EW - Elderly Waiver"+chr(9)+"EX - 65 and Older"+chr(9)+"LC - Long Term Care"+chr(9)+"MP - QMB/SLMB Only"+chr(9)+"N/A - No SSN"+chr(9)+"N/A - Not Applying"+chr(9)+"N/A - Not Deeming"+chr(9)+"N/A - PRIV"+chr(9)+"QI -QI"+chr(9)+"QW - QWD", avs_members_array(request_type_const, items)
                  DropListBox 235, y_pos - 5, 90, 15, "Select one..."+chr(9)+"N/A"+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_members_array(avs_status_const, items)
              End if
              EditBox 330, y_pos - 5, 50, 15, avs_members_array(avs_date_const, items)
              EditBox 390, y_pos - 5, 160, 15, avs_members_array(additional_info_const, items)
          Next
          y_pos = (80 + items * 15)
          Text 15, y_pos, 45, 15, "Other Notes:"
          EditBox 60, y_pos - 5, 225, 15, other_notes
          Text 290, y_pos, 60, 15, "Worker Signature:"
          EditBox 350, y_pos - 5, 120, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 475, (75 + (items * 15)), 40, 15
            CancelButton 515, (75 + (items * 15)), 40, 15
            PushButton 215, 0, 10, 15, "!", help_button_1
            PushButton 400, 20, 10, 15, "!", help_button_2
        EndDialog
    End if 
      'Member selection Dialog
    Do
        Do
            err_msg = ""
            Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
            cancel_confirmation
            If ButtonPressed = help_button_1 then
                tips_tricks_msg = MsgBox(help_button_text, vbInformation, "Tips and Tricks")
                err_msg = "LOOP" & err_msg
            End if
            If ButtonPressed = help_button_2 then
                tips_tricks_msg = MsgBox(help_button_2_text, vbInformation, "Tips and Tricks")
                err_msg = "LOOP" & err_msg
            End if

            'mandatory fields for all AVS_membs
            FOR items= 0 to UBound(avs_members_array, 2)
                'AVS Forms mandatory fields
                If initial_option = "AVS Forms" then
                    If avs_members_array(applicant_type_const, items) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the Applicant Type for: " & avs_members_array(member_info_const, items)
                    If avs_members_array(forms_status_const, items) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the forms status for: " & avs_members_array(member_info_const, items)
                    If avs_members_array(forms_status_const, items) = "Received - Incomplete" and trim(avs_members_array(additional_info_const, item)) = "" then err_msg = err_msg & vbcr & "* Enter the reason the AVS form is incomplete for: " & avs_members_array(member_info_const, items) & " in the 'additional information' field."
                End if
                'AVS Submission/Results mandatory fields
                If initial_option = "AVS Submission/Results" then
                    If avs_members_array(request_type_const, items) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the request type for: " & avs_members_array(member_info_const, item)
                    If avs_members_array(avs_status_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the request type for: " & avs_members_array(member_info_const, items)
                    If avs_members_array(avs_status_const, item) = "N/A" then
                        If left(avs_members_array(request_type_const, item), 3) <> "N/A" then err_msg = err_msg & vbcr & "* N/A is only a valid AVS status or process if the request type is also N/A."
                        If avs_members_array(additional_info_const, item) = "" then err_msg = err_msg & vbcr & "* Enter reason that N/A is the AVS Status or processed selected."
                    End if
                End if
                If trim(avs_members_array(avs_date_const, item)) = "" or isdate(avs_members_array(avs_date_const, item)) = False then err_msg = err_msg & vbcr & "* Enter the " & dialog_text & " status date for: " & avs_members_array(member_info_const, item)
            NEXT
            If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Enter your worker signature."
            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    '----------------------------------------------------------------------------------------------------ASSET DIALOG: for AVS Submission/Results members who have returned AVS results.
    For items= 0 to ubound(avs_members_array, 2)
        If avs_members_array(avs_status_const, item) = "Review Results" or avs_members_array(avs_status_const, item) = "Results After Decision" then
            STATS_counter = STATS_counter + 1
            'avs results information if any members meets review results or results after decision
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 351, 110, "Asset Information Dialog"
                GroupBox 5, 5, 340, 80, avs_members_array(avs_status_const, item) & " for "  & avs_members_array(member_info_const, item) & ":"
                Text 10, 25, 75, 10, "All accounts verified?"
                DropListBox 90, 20, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(accounts_verified_const, item)
                'Review Results selections
                If avs_members_array(avs_status_const, item) = "Review Results" then
                    Text 160, 25, 95, 10, "AVS Case Status for Member:"
                    DropListBox 260, 20, 80, 15, "Select one..."+chr(9)+"Close/Withdrawn"+chr(9)+"Eligible"+chr(9)+"Ineligible"+chr(9)+"N/A"+chr(9)+"Review in Progress"+chr(9)+"Transfer Penalty", avs_members_array(avs_results_const, item)
                'Results After Decision selections
                Elseif avs_members_array(avs_status_const, item) = "Results After Decision" then
                    Text 155, 25, 120, 10, "Accts after decision cleared in AVS?"
                    DropListBox 275, 20, 65, 12, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(avs_results_const, item)
                End if
                Text 10, 45, 75, 10, "Unreported accounts?"
                DropListBox 90, 40, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(unreported_assets_const, item)
                Text 170, 45, 100, 10, "AVS Report submitted to ECF?"
                DropListBox 275, 40, 65, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(ECF_const, item)
                Text 15, 65, 40, 10, "Asset notes:"
                EditBox 60, 60, 280, 15, avs_members_array(avs_returned_notes_const, item)
                Text 10, 95, 45, 10, "Other Notes:"
                EditBox 55, 90, 200, 15, other_notes
                ButtonGroup ButtonPressed
                'PushButton 220, 140, 30, 10, "Back", back_button
                OkButton 260, 90, 40, 15
                CancelButton 305, 90, 40, 15
            EndDialog

            Do
                Do
                    err_msg = ""
                    Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
                    cancel_confirmation
                    IF avs_members_array(accounts_verified_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Have all accounts been verified?"
                    If avs_members_array(unreported_assets_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Were there unreported accounts?"
                    If avs_members_array(accounts_verified_const, item) = "No" AND trim(avs_members_array(avs_returned_notes_const, item) = "") then err_msg = err_msg & vbcr & "* Explain answering 'No' to all accounts verified in the 'asset notes' field."
                    If avs_members_array(unreported_assets_const, item) = "Yes" AND trim(avs_members_array(avs_returned_notes_const, item) = "") then err_msg = err_msg & vbcr & "* Explain answering 'Yes' to unreported asset in the 'asset notes' field."
                    If avs_members_array(ECF_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Was the AVS report submitted to ECF for the case file?"
                    'Review Results option
                    If avs_members_array(avs_status_const, item) = "Review Results" then
                        If avs_members_array(avs_results_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the AVS Case Status for the member."
                        If avs_members_array(avs_results_const, item) = "N/A" and trim(avs_members_array(avs_returned_notes_const, item) = "") then err_msg = err_msg & vbcr & "* Enter the reason for the AVS Case Status was marked N/A."
                        If avs_members_array(ECF_const, item) = "No" then
                            If avs_members_array(avs_results_const, item) = "Close/Withdrawn" or avs_members_array(avs_results_const, item) = "Eligible" or avs_members_array(avs_results_const, item) = "Transfer Penalty" then err_msg = err_msg & vbcr & "* AVS Reports must be submitted to ECF unless the AVS status is N/A or Results in Progress."
                        End if
                    End if
                    'Results after decision options
                    IF avs_members_array(avs_status_const, item) = "Results After Decision" then
                        If avs_members_array(avs_results_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Have accounts after decision in AVS been cleared?"
                        If avs_members_array(avs_results_const, item) = "No" AND trim(avs_members_array(avs_returned_notes_const, item) = "") then err_msg = err_msg & vbcr & "* Explain answering 'No' to Accts after decision cleared in AVS in the 'asset notes' field."
                    End if
                    IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
                LOOP UNTIL err_msg = ""
                Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
            Loop until are_we_passworded_out = false					'loops until user passwords back in
        End if
    Next

    '----------------------------------------------------------------------------------------------------Conditional statements for actions and case noting
    'Determining if TIKL and verif request form info will be created - This will happen if initial_option = "AVS Forms" AND forms_status_const = "Initial Request"
    set_form_TIKL = False
    verif_request = False
    For i = 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, i) = "Initial Request" then
            set_form_TIKL = True
            verif_request = True
            Exit For
        End if
    Next

    'Determining if TIKL and verif request form info will be created - This will happen if initial_option = "AVS Forms" AND forms_status_const = "Initial Request"
    set_AVS_TIKL = False
    If initial_option = "AVS Submission/Results" then
        For i = 0 to ubound(avs_members_array, 2)
            If avs_members_array(avs_status_const, i) = "Submitting a Request" then
                set_AVS_TIKL = True
                Exit for
            End if
        Next
    End if

    'giving users the option to create another TIKL if the initial forms are incomplete.
    For i = 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, i) = "Received - Incomplete" then
            TIKL_msgbox = msgbox("Will you be sending another verification request for a completed AVS form?" & vbcr & vbcr & "Selecting YES will create another 10-day TIKL and case note that a verification request is being sent.", vbQuestion + vbYesNo, "Set another TIKL?")
            If TIKL_msgbox = vbYes then
                set_another_TIKL = True
                verif_request = True
            End if
            If TIKL_msgbox = vbNo then set_another_TIKL = False
            Exit for
        End if
    Next

    'Sending verif request for un reported assets found in the AVS system for the "AVS Submission/Results" option
    set_asset_TIKL = False
    If initial_option = "AVS Submission/Results" then
        For i = 0 to ubound(avs_members_array, 2)
            If avs_members_array(unreported_assets_const, i) = "Yes" then
                set_asset_TIKL = True
                Exit for
            End if
        Next
    End if

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    If set_form_TIKL = True then Call create_TIKL("DHS-7823 - AVS Auth Form(s) have been requested for this case. Review case file/notes, and take applicable actions.", 10, date, False, TIKL_note_text)
    If set_AVS_TIKL = True then Call create_TIKL("AVS 10-day check is due.", 11, date, False, TIKL_note_text)
    If set_another_TIKL = True then Call create_TIKL("An updated DHS-7823 - AVS Auth Form(s) has been requested for this case. Review case file/notes, and take applicable actions.", 10, date, False, TIKL_note_text)
    If set_asset_TIKL = True then Call create_TIKL("AVS unreported asset verification requested for the case. Review case file/notes, and take applicable actions.", 10, date, False, TIKL_note_text)

    'Adding closing message if the "AVS Submission/Results" option is selected and the AVS status is results after decision.
    If initial_option = "AVS Submission/Results" then
        For i = 0 to ubound(avs_members_array, 2)
            If avs_members_array(avs_status_const, i) = "Results After Decision" then closing_msg = closing_msg & vbcr & "Please review eligibility results to see if health care eligibility needs to be redetermined."
        Next
    End if

    '----------------------------------------------------------------------------------------------------The case note
    'Information for the case note
    If resize_counter = 0 then 'custom header  for single person cases
        If initial_option = "AVS Forms" then case_note_header = "--AVS Forms " & avs_members_array(forms_status_const, 0) & " for " & HC_process & "--"
        If initial_option = "AVS Submission/Results" then case_note_header = "--AVS System Request " & avs_members_array(avs_status_const, 0) & " for " & HC_process & "--"
    Else
        'generic header if more than one member case noting
        If initial_option = "AVS Forms" then case_note_header = "--AVS Forms for " & HC_process & " Information--"
        If initial_option = "AVS Submission/Results" then case_note_header = "--AVS System Request for " & avs_members_array(forms_status_const, 0) & " for " & HC_process & " Information--"
    End if

    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE(case_note_header)

    Call write_variable_in_CASE_NOTE("The following info is in regards to AVS members required to sign AVS Forms:")
    Call write_variable_in_CASE_NOTE("-----")
    'AVS member array output
    For items= 0 to ubound(avs_members_array, 2)
        Call write_bullet_and_variable_in_CASE_NOTE("Name", avs_members_array(member_info_const, item))
        'AVS Forms selections
        If initial_option = "AVS Forms" then
            Call write_bullet_and_variable_in_CASE_NOTE("Applicant Type", avs_members_array(applicant_type_const, item))
            Call write_bullet_and_variable_in_CASE_NOTE("AVS Form Status", avs_members_array(forms_status_const, item))
            'text for next case note variable
            If avs_members_array(forms_status_const, item) = "Initial Request" then
                forms_text = "Sent"
            Elseif avs_members_array(forms_status_const, item) = "Not Received" then
                forms_text = "Status"
            Else
                forms_text = "Rec'd"
            End if
            Call write_bullet_and_variable_in_CASE_NOTE("AVS Forms " & forms_text & " Date", avs_members_array(avs_date_const, item))
        End if
        'AVS Submission/Results selections
        If initial_option = "AVS Submission/Results" then
            Call write_bullet_and_variable_in_CASE_NOTE("Request Type", avs_members_array(request_type_const, item))
            Call write_bullet_and_variable_in_CASE_NOTE("AVS Status", avs_members_array(avs_status_const, item))
            'text for next dates for the case note variable
            If avs_members_array(avs_status_const, item) = "Submitting a Request" then
                status_text = "Sent"
            Else
                status_text = "Reviewed"
            End if
            Call write_bullet_and_variable_in_CASE_NOTE("AVS " & status_text & " Date", avs_members_array(avs_date_const, item))
        End if
        Call write_bullet_and_variable_in_CASE_NOTE("Additional Information", avs_members_array(additional_info_const, item))
        'Asset Dialog Information
        If avs_members_array(avs_status_const, item) = "Review Results" or avs_members_array(avs_status_const, item) = "Results After Decision" then
            Call write_bullet_and_variable_in_CASE_NOTE ("All Accounts Verified", avs_members_array(accounts_verified_const, item))
            Call write_bullet_and_variable_in_CASE_NOTE ("Unreported Assets", avs_members_array(unreported_assets_const, item))
            If avs_members_array(avs_status_const, item) = "Review Results" then
                Call write_bullet_and_variable_in_CASE_NOTE ("AVS Case Status for Member", avs_members_array(avs_results_const, item))
            Elseif avs_members_array(avs_status_const, item) = "Results After Decision" then
                Call write_bullet_and_variable_in_CASE_NOTE ("Accts after decision cleared in AVS?", avs_members_array(avs_results_const, item))
            End if
            Call write_bullet_and_variable_in_CASE_NOTE ("AVS Report Submitted to case file.", avs_members_array(ECF_const, item))
            Call write_bullet_and_variable_in_CASE_NOTE ("Asset Notes", avs_members_array(avs_returned_notes_const, item))
            Call write_variable_in_CASE_NOTE("-----")
        End if
    Next

    If verif_request = True then Call write_variable_in_CASE_NOTE("* Verification request sent to resident/AREP.")
    If set_form_TIKL = True then Call write_variable_in_case_note(TIKL_note_text)
    If set_AVS_TIKL = True then Call write_variable_in_CASE_NOTE("* AVS request submitted in AVS portal. TIKL set for 10 day check.") 'This is the verbiage for the case note from the HSR manual.
    If set_another_TIKL = true then
        Call write_variable_in_case_note(TIKL_note_text)
        Call write_variable_in_CASE_NOTE("* Sent verification request to complete the incomplete AVS form in case file.")
    End if
    If set_asset_TIKL = True then
        Call write_variable_in_case_note(TIKL_note_text)
        Call write_variable_in_CASE_NOTE("* Sent verification request for unreported assets in case file.")
    End if
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    'Providing the option to run the avs option
    For items= 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, item) = "Received - Complete" then
            confirm_msgbox = msgbox("Do you wish to case note submitting the AVS request?", vbQuestion + vbYesNo, "Submit the AVS Request?")
            If confirm_msgbox = vbNo then
                run_initial_option = False
            elseif confirm_msgbox = vbYes then
                PF3 ' to save case note
                run_initial_option = True
                initial_option = "AVS Submission/Results"
            End if
            exit for
        End if
    Next
    If run_initial_option = False then exit do
Loop

If verif_request = True then
    'Outputting AVS verbiage to Word
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set objDoc = objWord.Documents.Add()
    objWord.Caption = "AVS Verification Request Verbiage"
    Set objSelection = objWord.Selection
    objSelection.PageSetup.LeftMargin = 50
    objSelection.PageSetup.RightMargin = 50
    objSelection.PageSetup.TopMargin = 30
    objSelection.PageSetup.BottomMargin = 25
    objSelection.ParagraphFormat.SpaceAfter = 0
    objSelection.Font.Name = "Calibri"

    objSelection.Font.Size = "14"
    objSelection.Font.Bold = False
    objSelection.TypeText "All applicants (or their authorized representative) must sign and return the enclosed Authorization(s) to Obtain Financial Information from AVS. " & vbCr
    objSelection.TypeText "Each applicant needs their own form. The form is mandatory to determine eligibility for certain health care programs. " & vbCr
    objSelection.TypeText "Spouses who live together must sign each other's forms. If you are a sponsored immigrant, your sponsor(s) and sponsor(s)' spouse(s) must sign your form."

    'closing message with reminder to send to ECF and Word Document
    closing_msg = closing_msg & vbcr & "Remember to send your verification request in ECF with the correct verbiage." & vbcr & vbcr & "This verbiage has been outputted to a Word document for your convenience."
End if

STATS_counter = STATS_counter - 1   'removing increment as we start with 1.
script_end_procedure_with_error_report(closing_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/30/2022
'--Tab orders reviewed & confirmed----------------------------------------------12/30/2022
'--Mandatory fields all present & Reviewed--------------------------------------12/30/2022
'--All variables in dialog match mandatory fields-------------------------------12/30/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/30/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------12/30/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------12/30/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-12/30/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/30/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------12/30/2022
'--PRIV Case handling reviewed -------------------------------------------------12/30/2022
'--Out-of-County handling reviewed----------------------------------------------12/30/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/30/2022
'--BULK - review output of statistics and run time/count (if applicable)--------12/30/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---12/30/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/30/2022
'--Incrementors reviewed (if necessary)-----------------------------------------12/30/2022
'--Denomination reviewed -------------------------------------------------------12/30/2022
'--Script name reviewed---------------------------------------------------------12/30/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------12/30/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------12/30/2022
'--comment Code-----------------------------------------------------------------12/30/2022
'--Update Changelog for release/update------------------------------------------12/30/2022
'--Remove testing message boxes-------------------------------------------------12/30/2022
'--Remove testing code/unnecessary code-----------------------------------------12/30/2022
'--Review/update SharePoint instructions----------------------------------------12/30/2022----------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------12/30/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------12/30/2022
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------12/30/2022
'--Complete misc. documentation (if applicable)---------------------------------12/30/2022
'--Update project team/issue contact (if applicable)----------------------------12/30/2022
