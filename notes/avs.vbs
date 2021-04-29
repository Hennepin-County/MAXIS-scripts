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
call changelog_update("03/23/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'TODO: Carry over other notes area into asset dialog
'TODO: Out of county case handling
'TODO: Figure out back and next buttons functionality
'TODO: Start instructions
'TODO: Time study
'TODO: stats increment counter
'TODO: Comment code
'TODO: Hot Topics

'----------------------------------------------------------------------------------------------------The script
closing_msg = "Success! Your AVS case note has been created. Please review for accuracy & any additional information."

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

msgbox "Oh hi there! 1"
'----------------------------------------------------------------------------------------------------Initial dialog
initial_help_text = "*** What is the AVS? ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & _
"The Account Validation Service (AVS) is a web-based service that provides information about some accounts held in financial institutions. It does not provide information on property assets such as cars or homes. AVS must be used once at application, and when a person changes to a Medical Assistance for People Who Are Age 65 or Older and People Who Are Blind or Have a Disability (MA-ABD) basis of eligibility and are subject to an asset test." & vbNewLine & vbNewLine & _
"If a resident is already on a MHCP with an asset test or a MHCP with an asset test isn't being applied for then the AVS should not be run. This verification is not meant for any other public assitance programs besides health care."

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 85, "AVS Initial Selection Dialog"
  EditBox 75, 10, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
  PushButton 135, 10, 10, 15, "!", initial_help_button
  DropListBox 75, 30, 105, 15, "Select one..."+chr(9)+"AVS Forms"+chr(9)+"AVS Submission/Results", initial_option
  DropListBox 75, 45, 105, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Renewal", HC_process
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
            tips_tricks_msg = MsgBox(initial_help_text, vbInformation, "Tips and Tricks") 'see help_button_text above for details of the text
            err_msg = "LOOP" & err_msg
        End if
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid MAXIS Case Number."
        If initial_option = "Select one..." then err_msg = err_msg & vbcr & "* Select the AVS process."
        If HC_process = "Select one..." then err_msg = err_msg & vbcr & "* Select the health care process."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If HC_process = "Renewal" then script_end_procedure("The AVS system is not required at renewals at this time. The script will now end.")

MAXIS_background_check
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privilged, and you do not have access. The script will now end.")

' 'Adding supports for users who are using the AVS incorrectly to verify assets for other programs besides HC. 
' EmReadscreen HC_status, 4, 12, 74
' If HC_status = "PEND" or HC_status = "ACTV" or HC_status = "REIN" then
'     HC_pending = True
' Else
'     script_end_procedure("This case is not acitve, pending or in reinstatement status. An AVS should only be run on acitve or pending health care cases with an asset test. The script will now end.")
' End if

'----------------------------------------------------------------------------------------------------Gathering the member/AREP/Sponsor information for signature selection array
Call navigate_to_MAXIS_screen("STAT", "MEMB")
'Setting up main array
avs_membs = 0
Dim avs_members_array()
ReDim avs_members_array(additional_info_const, 0)

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
const avs_returned_no_const    = 11
const avs_date_const           = 12
const accounts_verified_const  = 13
const unreported_assets_const  = 14
const ECF_const                = 15
const additional_info_const    = 16

add_to_array = False    'defaulting to false
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EmReadscreen relationship_code, 2, 10, 42
    EmReadscreen client_age, 3, 8, 76
    EmReadscreen client_ssn, 11, 7, 42
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")

    If relationship_code = "01" then add_to_array = True    'applicants of HC yes
    If relationship_code = "02" then add_to_array = True    'spouses of HC applicants yes

    If client_ssn = "___ __ ____" then
        client_ssn = ""
        'Folks who have no ssn are not required to submit an AVS inquiry
        If initial_option = "AVS Submission/Results" then avs_members_array(request_type_const, avs_membs) = "N/A - No SSN"
    Else
        client_ssn = replace(client_ssn, " ", "")
    End if

    If trim(client_age) < "21" then add_to_array = False  'under 21 are not required to sign per EPM 2.3.3.2.1 Asset Limits

    If add_to_array = True then
        ReDim Preserve avs_members_array(additional_info_const, avs_membs)
        avs_members_array(member_info_const,    avs_membs) = ref_nbr & " " & last_name & first_name
        avs_members_array(member_number_const,  avs_membs) = ref_nbr
        avs_members_array(member_name_const,    avs_membs) = first_name & "" & last_name
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_membs = avs_membs + 1
    End if
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'Handling in case no members are idenified as needing the form. Helping to reduce errors for workers.
If avs_membs = 0 then script_end_procedure_with_error_report("No members on this case are required by policy to sign the AVS form. Please review case if necessary.")

Call navigate_to_MAXIS_screen("STAT", "MEMI")
For item = 0 to Ubound(avs_members_array, 2)
    EmWriteScreen avs_members_array(member_number_const, item), 20, 76
    transmit
    EmReadscreen marital_status, 1, 7, 40
    'msgbox marital_status
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

Call navigate_to_MAXIS_screen("STAT", "TYPE")
For item = 0 to Ubound(avs_members_array, 2)
    If avs_members_array(hc_applicant_const, item) = "" then
        'msgbox avs_members_array(member_name_const, item)
        'adding TYPE Information to be output into the dialog and case note
        row = 6
        Do
            EmReadscreen type_memb_number, 2, row, 3
            EmReadscreen applicant_type, 1, row, 37
            'msgbox  "row: " & row & vbcr & _
            '        "memb #: " & type_memb_number & vbcr & _
            '        "hc type: " & applicant_type
            If type_memb_number = avs_members_array(member_number_const, item) then
                If applicant_type = "Y" then
                    'msgbox "match"
                    avs_members_array(hc_applicant_const, item) = True
                    If avs_members_array(marital_status_const, item) = "M" then
                        avs_members_array(applicant_type_const, item) = "Applying/Spouse"
                    Else
                        avs_members_array(applicant_type_const,      item) = "Applying"
                    End if
                    exit do
                Elseif applicant_type = "N" then
                    avs_members_array(hc_applicant_const, item) = False
                    If avs_members_array(marital_status_const, item) = "M" then
                        avs_members_array(applicant_type_const, item) = "Spouse"
                    Else
                        avs_members_array(applicant_type_const,      item) = "Not Applying"
                    End if
                    exit do
                End if
            Else
                row = row + 1
            End if
        Loop until trim(type_memb_number) = ""
	End if
Next

'For item = 0 to Ubound(avs_members_array, 2)
'    If avs_members_array(hc_applicant_const, item) = True then
'        msgbox "applicant: " & avs_members_array(member_name_const, item) & vbcr & "HC type: " & avs_members_array(applicant_type_const, item)
'    Else
'        msgbox "applicant: " & avs_members_array(member_name_const, item) & vbcr & "HC type: " & avs_members_array(applicant_type_const, item)
'    End if
'Next

Do
    If confirm_msgbox = vbYes then
        For item = 0 to ubound(avs_members_array, 2)
             avs_members_array(forms_status_const,     item) = ""
            avs_members_array(avs_status_const,        item) = ""
            avs_members_array(request_type_const,      item) = ""
            avs_members_array(avs_results_const,       item) = ""
            avs_members_array(avs_returned_no_const,   item) = ""
            avs_members_array(avs_date_const,          item) = ""
            avs_members_array(accounts_verified_const, item) = ""
            avs_members_array(unreported_assets_const, item) = ""
            avs_members_array(ECF_const,               item) = ""
            avs_members_array(additional_info_const,   item) = ""
        Next
    End if
    '----------------------------------------------------------------------------------------------------SELECTING AVS MEMBERS: Based on who is required to sign form/submit AVS
    'Text for the next dialogs based on initial option selected by the user
    If initial_option = "AVS Forms" then
        selection_text = "Select all members REQUIRED to sign AVS form(s):"
        type_text = "Applicant"
        dialog_text = "Forms"
        help_button_text = "*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Information source: DHS-7823 Form - Authorization to Obtain Financial Information from the Account Validation Service (AVS)." & vbcr & vbcr & _
        "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _
        "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _
        "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country."

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

        help_button_text = "*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Information source: DHS-7823 Form - Authorization to Obtain Financial Information from the Account Validation Service (AVS)." & vbcr & vbcr & _
        "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _
        "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _
        "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country."

        help_button_2_text = "*** What date should I enter here? ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & _
        "The AVS Status will determine what you will enter in this field. These will usually be the current date." & vbNewLine & vbNewLine & _
        "- Submitting a Request: Enter the date the request was sent in the AVS system." & vbNewLine & _
        "- Review Results or Results After Decision: Enter the date the results were reviewed in the AVS system."
    End if

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 200, (50 + (item * 20)), "AVS Member Selection Dialog"
        Text 5, 5, 180, 10, selection_text
        ButtonGroup ButtonPressed
        PushButton 170, 0, 10, 15, "!", help_button
        For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 1 then checkbox 15, (20 + (item * 20)), 100, 15, avs_members_array(member_info_const, item), avs_members_array(checked_const, item)
        Next
        ButtonGroup ButtonPressed
        OkButton 85, (30 + (item * 20)), 45, 15
        CancelButton 135, (30 + (item * 20)), 45, 15
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
            FOR item = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
                If avs_members_array(checked_const, item) = 1 then checked_count = checked_count + 1 'Ignores and blank scanned in persons/strings to avoid a blank checkbox
            NEXT
            'msgbox "checked count: " & checked_count
            If checked_count = 0 then err_msg = err_msg & vbcr & "* Select all persons responsible for signing the AVS form."
            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved

        'TODO: Is this necessary anymore?
        'Revaluing the checked or selected names based on the user selection in the dialog
        FOR item = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 0 then avs_members_array(checked_const, item) = 0
        NEXT

        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    '----------------------------------------------------------------------------------------------------Adding in information about the AVS Members selected
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 575, (115 + (avs_membs * 15)), "AVS Member Information Dialog"
      GroupBox 10, 5, 550, (60 + (avs_membs * 15)), "Complete the following information for required AVS members:"
      ButtonGroup ButtonPressed
        PushButton 215, 0, 10, 15, "!", help_button_1
      Text 20, 25, 520, 10, "----------AVS Member--------------------------------" & type_text & " Type---------------------------" & dialog_text & " Status-------------------" & dialog_text & " Sent/Rec'd Date-------------------Additional Information----------------"
      ButtonGroup ButtonPressed
        PushButton 400, 20, 10, 15, "!", help_button_2
        For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 1 then
                y_pos = (50 + item * 20)
                Text 20, y_pos, 110, 15, avs_members_array(member_info_const, item)
                If initial_option = "AVS Forms" then
                    DropListBox 130, y_pos - 5, 70, 15, "Select one..."+chr(9)+"Applying"+chr(9)+"Applying/Spouse"+chr(9)+"Deeming"+chr(9)+"Not Applying"+chr(9)+"Spouse", avs_members_array(applicant_type_const, item)
                    DropListBox 225, y_pos - 5, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", avs_members_array(forms_status_const, item)
                End if

                If initial_option = "AVS Submission/Results" then
                        DropListBox 120, y_pos - 5, 90, 15, "Select one..."+chr(9)+"BI - Brain Injury Waiver"+chr(9)+"BX - Blind"+chr(9)+"CA - CAC Waiver"+chr(9)+"CD - CADI Waiver"+chr(9)+"DD - DD Waiver"+chr(9)+"DP - MA-EPD"+chr(9)+"DX - Disability"+chr(9)+"EH - EMA"+chr(9)+"EW - Elderly Waiver"+chr(9)+"EX - 65 and Older"+chr(9)+"LC - Long Term Care"+chr(9)+"MP - QMB/SLMB Only"+chr(9)+"N/A - No SSN"+chr(9)+"N/A - Not Applying"+chr(9)+"N/A - Not Deeming"+chr(9)+"N/A - PRIV"+chr(9)+"QI - QI"+chr(9)+"QW     - QWD", avs_members_array(request_type_const, item)
                    DropListBox 225, y_pos - 5, 90, 15, "Select one..."+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_members_array(avs_status_const, item)
                End if
                EditBox 330, y_pos - 5, 50, 15, avs_members_array(avs_date_const, item)
                EditBox 390, y_pos - 5, 160, 15, avs_members_array(additional_info_const, item)
            End if
        Next
        y_pos = (80 + item * 15)
        Text 15, y_pos, 45, 15, "Other Notes:"
        EditBox 60, y_pos - 5, 225, 15, other_notes
        Text 290, y_pos, 60, 15, "Worker Signature:"
        EditBox 350, y_pos - 5, 120, 15, worker_signature
        ButtonGroup ButtonPressed
          OkButton 475, (75 + (item * 15)), 40, 15
          CancelButton 515, (75 + (item * 15)), 40, 15
      EndDialog

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
            FOR item = 0 to UBound(avs_members_array, 2)
                If initial_option = "AVS Forms" then
                    If avs_members_array(applicant_type_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the Applicant Type for: " & avs_members_array(member_info_const, item)
                    If avs_members_array(forms_status_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the forms status for: " & avs_members_array(member_info_const, item)
                    If (avs_members_array(forms_status_const, item) = "Received - Incomplete" and trim(avs_members_array(additional_info_const, item)) = "") then err_msg = err_msg & vbcr & "* Enter the reason the AVS form is incomplete for: " & avs_members_array(member_info_const, item) & " in the 'additional information' field."
                End if
                If initial_option = "AVS Submission/Results" then
                    If avs_members_array(request_type_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the request type for: " & avs_members_array(member_info_const, item)
                    If avs_members_array(avs_status_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Enter the request type for: " & avs_members_array(member_info_const, item)
                End if
                If trim(avs_members_array(avs_date_const, item)) = "" or isdate(avs_members_array(avs_date_const, item)) = False then err_msg = err_msg & vbcr & "* Enter the " & dialog_text & " status date for: " & avs_members_array(member_info_const, item)
            NEXT
            If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Enter your worker signature."
            IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    For item = 0 to ubound(avs_members_array, 2)
        If avs_members_array(avs_status_const, item) = "Review Results" or avs_members_array(avs_status_const, item) = "Results After Decision" then
            'avs results information if any mambers meets review resutls or results after decision
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 351, 110, "Asset Information Dialog"
                GroupBox 5, 5, 340, 80, avs_members_array(avs_status_const, item) & " for "  & avs_members_array(member_info_const, item) & ":"
                Text 10, 25, 75, 10, "All accounts verified?"
                DropListBox 90, 20, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(accounts_verified_const, item)
                If avs_members_array(avs_status_const, item) = "Review Results" then
                    Text 160, 25, 95, 10, "AVS Case Status for Member:"
                    DropListBox 260, 20, 80, 15, "Select one..."+chr(9)+"Close/Withdrawn"+chr(9)+"Eligible"+chr(9)+"Ineligible"+chr(9)+"Review in Progress"+chr(9)+"Transfer Penalty", avs_members_array(avs_results_const, item)
                Elseif avs_members_array(avs_status_const, item) = "Results After Decision" then
                    Text 155, 25, 120, 10, "Accts after decision cleared in AVS?"
                    DropListBox 275, 20, 65, 12, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(avs_results_const, item)
                End if
                Text 10, 45, 75, 10, "Unreported accounts?"
                DropListBox 90, 40, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(unreported_assets_const, item)
                Text 170, 45, 100, 10, "AVS Report submitted to ECF?"
                DropListBox 275, 40, 65, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", avs_members_array(ECF_const, item)
                Text 15, 65, 40, 10, "Asset notes:"
                EditBox 60, 60, 280, 15, avs_members_array(avs_returned_no_const, item)
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
                    'error message variable based on option
                    If avs_members_array(avs_results_const, item) = "Select one..." then
                        If avs_members_array(avs_status_const, item) = "Review Results" then
                            err_msg = err_msg & vbcr & "* Enter the AVS Case Status for the member."
                        Elseif avs_members_array(avs_status_const, item) = "Results After Decision" then
                            err_msg = err_msg & vbcr & "* Have accounts after decicion in AVS been cleared?"
                        End if
                    End if
                    If avs_members_array(ECF_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Was the AVS report submitted to ECF for the case file?"
                    If (avs_members_array(avs_results_const, item) = "No" or avs_members_array(accounts_verified_const, item) = "No") AND _
                    trim(avs_members_array(avs_returned_no_const, item) = "") then err_msg = err_msg & vbcr & "* Explain answering 'No' to one or more questions in the dialog in the 'asset notes' field."
                    If avs_members_array(unreported_assets_const, item) = "Yes" AND trim(avs_members_array(avs_returned_no_const, item) = "") then err_msg = err_msg & vbcr & "* Explain answering 'Yes' to unreported asset in the 'asset notes' field."

                    IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
                LOOP UNTIL err_msg = ""
                Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
            Loop until are_we_passworded_out = false					'loops until user passwords back in
        End if
    Next

    '----------------------------------------------------------------------------------------------------Conditional statements for actions and case noting
    'Determining if TIKL and verif request form info will be created - This will happen if initial_option = "AVS Forms" AND forms_status_const = "Initial Request"
    Set_TIKL = False
    verif_request = False
    For i = 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, i) = "Initial Request" then
            Set_TIKL = True
            verif_request = True
            Exit For
        End if
    Next

    'giving users the option to create another TIKL if the initial forms are incomplete.
    For i = 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, i) = "Received - Incomplete" then
            TIKL_msgbox = msgbox("Do you wish to send another verification requesting a completed AVS form?" & vbcr & vbcr & "Selecting YES will create another 10-day TIKL and case note that a verification request is being sent.", vbQuestion + vbYesNo, "Set another TIKL?")
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

    'set_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    If set_TIKL = True then Call create_TIKL("DHS-7823 - AVS Auth Form(s) have been requested for this case. Review ECF and case notes, and take applicable actions.", 10, date, False, TIKL_note_text)
    If set_another_TIKL = True then Call create_TIKL("An updated DHS-7823 - AVS Auth Form(s) has been requested for this case. Review ECF and case notes, and take applicable actions.", 10, date, False, TIKL_note_text)
    If set_asset_TIKL = True then Call create_TIKL("AVS unreported asset verification requested for the case. Review ECF and case notes, and take applicable actions.", 10, date, False, TIKL_note_text)

    'Adding closing message if the "AVS Submission/Results" option is selected and the AVS status is results after decision.
    If initial_option = "AVS Submission/Results" then
        For i = 0 to ubound(avs_members_array, 2)
            If avs_members_array(avs_status_const, i) = "Results After Decision" then closing_msg = closing_msg & vbcr & "Please review eligibility results to see if health care eligibility needs to be redetermined."
        Next
    End if

    '----------------------------------------------------------------------------------------------------The case note
    'Information for the case note
    If initial_option = "AVS Forms" then case_note_header = "--AVS Forms for " & HC_process & " Information--"
    If initial_option = "AVS Submission/Results" then case_note_header = "--AVS System Request for " & HC_process & " Information--"

    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE(case_note_header)

    Call write_variable_in_CASE_NOTE("The following information is in regards to AVS members required to sign the AVS Forms:")
    Call write_variable_in_CASE_NOTE("-----")
    'HH member array output
    For item = 0 to ubound(avs_members_array, 2)
        If avs_members_array(checked_const, item) = 1 then
            Call write_bullet_and_variable_in_CASE_NOTE("Name", avs_members_array(member_info_const, item))

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

            If initial_option = "AVS Submission/Results" then
                Call write_bullet_and_variable_in_CASE_NOTE("Request Type", avs_members_array(request_type_const, item))
                Call write_bullet_and_variable_in_CASE_NOTE("AVS Status", avs_members_array(avs_status_const, item))
                'text for next case note variable
                If avs_members_array(avs_status_const, item) = "Submitting a Request" then
                    status_text = "Sent"
                Else
                    status_text = "Reviewed"
                End if
                Call write_bullet_and_variable_in_CASE_NOTE("AVS " & status_text & " Date", avs_members_array(avs_date_const, item))
            End if
            Call write_bullet_and_variable_in_CASE_NOTE("Additional Information", avs_members_array(additional_info_const, item))

            If avs_members_array(avs_status_const, item) = "Review Results" or avs_members_array(avs_status_const, item) = "Results After Decision" then
                Call write_bullet_and_variable_in_CASE_NOTE ("All Accounts Verified", avs_members_array(accounts_verified_const, item))
                Call write_bullet_and_variable_in_CASE_NOTE ("Unreported Assets", avs_members_array(unreported_assets_const, item))
                If avs_members_array(avs_status_const, item) = "Review Results" then
                    Call write_bullet_and_variable_in_CASE_NOTE ("AVS Case Status for Member", avs_members_array(avs_results_const, item))
                Elseif avs_members_array(avs_status_const, item) = "Results After Decision" then
                    Call write_bullet_and_variable_in_CASE_NOTE ("Accts after decision cleared in AVS?", avs_members_array(avs_results_const, item))
                End if
                Call write_bullet_and_variable_in_CASE_NOTE ("AVS Report Submitted to ECF?", avs_members_array(ECF_const, item))
                Call write_bullet_and_variable_in_CASE_NOTE ("Asset Notes", avs_members_array(avs_returned_no_const, item))
                Call write_variable_in_CASE_NOTE("-----")
            End if
        End if
    Next

    If verif_request = True then Call write_variable_in_CASE_NOTE("* Verification request sent to via ECF.")
    If set_TIKL = true then Call write_variable_in_case_note(TIKL_note_text)
    If set_another_TIKL = true then
        Call write_variable_in_case_note(TIKL_note_text)
        Call write_variable_in_CASE_NOTE("* Sent verification request to complete the incomplete AVS form in ECF.")
    End if
    If set_asset_TIKL = True then
        Call write_variable_in_case_note(TIKL_note_text)
        Call write_variable_in_CASE_NOTE("* Sent verification request for unreported assets in ECF.")
    End if
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    'Providing the option to run the avs option
    For item = 0 to ubound(avs_members_array, 2)
        If avs_members_array(forms_status_const, item) = "Received - Complete" then
            confirm_msgbox = msgbox("Do you wish to case note submitting the AVS request?", vbQuestion + vbYesNo, "Submit the AVS Request?")
            If confirm_msgbox = vbNo then
                run_initial_option = False
            elseif confirm_msgbox = vbYes then
                PF3 ' to save case note
                run_initial_option = True
                initial_option = "AVS Submission/Results"
                msgbox initial_option
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

script_end_procedure_with_error_report(closing_msg)
