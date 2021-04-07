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

'----------------------------------------------------------------------------------------------------The script
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

initial_option = "AVS Forms"     'Testing code
HC_process = "Application"      'Testing code 
MAXIS_case_number = "299320"    'Testing code 

'----------------------------------------------------------------------------------------------------Initial dialog 
BeginDialog Dialog1, 0, 0, 166, 85, "AVS Initial Selection Dialog"
  EditBox 75, 10, 55, 15, MAXIS_case_number
  DropListBox 75, 30, 85, 15, "Select one..."+chr(9)+"AVS Forms"+chr(9)+"AVS Submission/Results", initial_option
  DropListBox 75, 45, 85, 15, "Select one..."+chr(9)+"Application", HC_process
    ' DropListBox 75, 45, 85, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change in Basis"+chr(9)+"Renewal", HC_process
  ButtonGroup ButtonPressed
    OkButton 75, 65, 40, 15
    CancelButton 120, 65, 40, 15
  Text 25, 15, 45, 10, "Case number:"
  Text 5, 35, 70, 10, "Select AVS Process:"
  Text 5, 50, 70, 10, "Select HC Process:"
EndDialog

'Initial dialog: user will input case number and initial options
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				    
		cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
		Call validate_MAXIS_case_number(err_msg, "*")
        If initial_option = "Select one..." then err_msg = err_msg & vbcr & "* Select the AVS process."
        If HC_process = "Select one..." then err_msg = err_msg & vbcr & "* Select the health care process."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privilged, and you do not have access. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member/AREP/Sponsor information for signature selection array 
'Setting up main array 
avs_membs = 0
Dim avs_members_array()
ReDim avs_members_array(9, 0)

const member_number_const   = 0
const member_info_const     = 1 
const member_name_const     = 2 
const marital_status_const  = 3
const checked_const         = 4
const hc_applicant_const    = 5
const applicant_type_const  = 6
const forms_status_const    = 7
const avs_date_const        = 8
const additional_info_const = 9

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
        'add_to_array = False    'No SSN's do not able/required to run AVS TODO: no SSN needs to sign not submit AVS
    Else 
        client_ssn = replace(client_ssn, " ", "")
    End if 
    
    If trim(client_age) < "18" then add_to_array = False  'under 18 are not required to sign 
    
    If add_to_array = True then   
        ReDim Preserve avs_members_array(9,     avs_membs)
        avs_members_array(member_info_const,    avs_membs) = ref_nbr & " " & last_name & first_name
        avs_members_array(member_number_const,  avs_membs) = ref_nbr
        avs_members_array(member_name_const,    avs_membs) = first_name & "" & last_name
        avs_members_array(marital_status_const, avs_membs) = ""
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_members_array(hc_applicant_const,   avs_membs) = ""         'defaulted to blank until determined if applicant or not (based on STAT/TYPE)
        avs_members_array(applicant_type_const, avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_members_array(forms_status_const,   avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_members_array(avs_date_const,       avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_members_array(additional_info_const,avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_membs = avs_membs + 1
    End if   
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'Handling in case no members are idenified as needing the form. Helping to reduce errors for workers. 
If avs_membs = 0 then script_end_procedure_with_error_report("No members on this case are required by policy to sign the AVS form. Please review case if necessary.")

Call navigate_to_MAXIS_screen("STAT", "MEMI")
For item = 0 to Ubound(avs_members_array, 2) 
    Call write_value_and_transmit(avs_members_array(member_number_const), 20, 76)
    EmReadscreen marital_status, 1, 7, 40
    msgbox marital_status
    avs_members_array(marital_status_const, item) = marital_status
Next 
    
'----------------------------------------------------------------------------------------------------AREP information if applicable to be added to array 
Call navigate_to_MAXIS_screen("STAT", "AREP")
EMReadScreen arep_name, 37, 4, 32
arep_name = replace(arep_name, "_", "")
If trim(arep_name) <> "" then  
    client_string = "AREP: " & trim(arep_name)
    client_array = client_array & trim(client_string) & "|"
    'Adding to main array 
    ReDim Preserve avs_members_array(9,     avs_membs)
    avs_members_array(member_info_const,    avs_membs) = "AREP: " & trim(arep_name)
    avs_members_array(member_number_const,  avs_membs) = ""
    avs_members_array(member_name_const,    avs_membs) = trim(arep_name)
    avs_members_array(marital_status_const, avs_membs) = ""
    avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
    avs_members_array(hc_applicant_const,   avs_membs) = FALSE      'AREP's will not be applicants          
    avs_members_array(applicant_type_const, avs_membs) = "Not Applying"
    avs_members_array(forms_status_const,   avs_membs) = ""  
    avs_members_array(avs_date_const,       avs_membs) = ""         'defaulted to blank until determined later in user dialog 
    avs_members_array(additional_info_const,avs_membs) = ""         'defaulted to blank until determined later in user dialog 
    avs_membs = avs_membs + 1
End if

'----------------------------------------------------------------------------------------------------SPONSOR information if applicable to be added to array 
Call navigate_to_MAXIS_screen("STAT", "SPON")
EmReadscreen total_spon_panels, 1, 2, 78
Do 
    If total_spon_panels = "0" then exit do 
    EMReadScreen spon_name, 20, 8, 38
    spon_name = replace(spon_name, "_", "")
    If trim(spon_name) <> "" then  
        'Adding to main array   
        ReDim Preserve avs_members_array(9,     avs_membs)
        avs_members_array(member_info_const,    avs_membs) = "Sponsor: " & trim(spon_name)
        avs_members_array(member_number_const,  avs_membs) = ""
        avs_members_array(member_name_const,    avs_membs) = trim(spon_name)
        avs_members_array(marital_status_const, avs_membs) = ""
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_members_array(hc_applicant_const,   avs_membs) = FALSE      'AREP's will not be applicants          
        avs_members_array(applicant_type_const, avs_membs) = "Deeming"
        avs_members_array(forms_status_const,   avs_membs) = ""  
        avs_members_array(avs_date_const,       avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_members_array(additional_info_const,avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_membs = avs_membs + 1
    End if
    transmit
    EMReadScreen last_panel, 5, 24, 2
Loop until last_panel = "ENTER"	'This means that there are no other faci panels

'msgbox "AVS Membs: " & avs_membs

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

For item = 0 to Ubound(avs_members_array, 2) 
    If avs_members_array(hc_applicant_const, item) = True then 
        msgbox "applicant: " & avs_members_array(member_name_const, item) & vbcr & "HC type: " & avs_members_array(applicant_type_const, item)
    Else
        msgbox "applicant: " & avs_members_array(member_name_const, item) & vbcr & "HC type: " & avs_members_array(applicant_type_const, item) 
    End if 
Next

'----------------------------------------------------------------------------------------------------SELECTING AVS MEMBERS: Based on who is required to sign form/submit AVS
'Text for the next dialogs based on initial option selected by the user
If initial_option = "AVS Forms" then 
    selection_text = "Select all members REQUIRED to sign AVS form(s):"
    dialog_text = "Forms"
    help_button_text = ""
    help_button_1_text = ""
    help_button_2_text = ""
End if 
If initial_option = "AVS Submission/Results" then 
    selection_text = "Select all members who meet AVS requirement:"
    dialog_text = "AVS"
    help_button_text = ""
    help_button_1_text = ""
    help_button_2_text = ""
End if 

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 200, 125, "AVS Member Selection Dialog"
    Text 5, 5, 180, 10, selection_text
    ButtonGroup ButtonPressed
    PushButton 170, 0, 10, 15, "!", help_button
    For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        If avs_members_array(checked_const, item) = 1 then checkbox 10, (20 + (item * 20)), 100, 10, avs_members_array(member_info_const, item), avs_members_array(checked_const, item)         
    Next 
    ButtonGroup ButtonPressed
    OkButton 85, 105, 45, 15
    CancelButton 135, 105, 45, 15
EndDialog

'Member selection Dialog 
Do 
    Do 
        err_msg = ""
        Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
        cancel_without_confirmation
        If ButtonPressed = help_button then 
            tips_tricks_msg = MsgBox("*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Information source: DHS-7823 Form - Authorization to Obtain Financial Information from the Account Validation Service (AVS)." & vbcr & vbcr & _
            "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _ 
            "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _ 
            "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country.", vbInformation, "Tips and Tricks")        
            err_msg = "LOOP" & err_msg
        End if 
        'ensuring that users have 
        checked_count = 0
        FOR item = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 1 then checked_count = checked_count + 1 'Ignores and blank scanned in persons/strings to avoid a blank checkbox
        NEXT
        msgbox "checked count: " & checked_count
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
BeginDialog Dialog1, 0, 0, 555, (115 + (avs_membs * 15)), "AVS Member Information Dialog"
  GroupBox 10, 5, 400, 115, "Complete the following information for required AVS members: "
  ButtonGroup ButtonPressed
    PushButton 215, 0, 10, 15, "!", help_button_1
  Text 20, 25, 515, 10, "----------AVS Member------------------------------Applicant Type---------------------------" & dialog_text &  "Status-------------------" & dialog_text & " Sent/Rec'd Date------------------------------Additional Information--------"
  ButtonGroup ButtonPressed
    PushButton 385, 20, 10, 15, "!", help__button_2
    For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        If avs_members_array(checked_const, item) = 1 then 
            Text 20, (40 + (item * 15)), 110, 10, avs_members_array(member_info_const, item)
            DropListBox 135, (20 + (item * 20)), 70, 10, "Select one..."+chr(9)+"Applying"+chr(9)+"Applying/Spouse"+chr(9)+"Deeming"+chr(9)+"Not Applying"+chr(9)+"Spouse", avs_members_array(applicant_type_const, item)
            If initial_option = "AVS Forms" then DropListBox 215, (20 + (item * 20)), 60, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", avs_members_array(forms_status_const, item)   
            'If initial_option = "AVS Submission/Results" then DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_members_array(avs_status_const, item)  
            EditBox 325, (20 + (item * 20)), 50, 15, avs_members_array(avs_date_const, item)
            EditBox 385, (20 + (item * 20)), 50, 15, avs_members_array(additional_info_const, item)
        End if 
    Next 
    
    Text 45, (80 + (item * 15)), 45, 10, "Other Notes:"
    EditBox 50, (75 + (item * 15)), 250, 15, other_notes
    Text 290, (80 + (item * 15)), 60, 10, "Worker Signature:"
    EditBox 350, (75 + (item * 15)), 150, 15, worker_signature
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
            tips_tricks_msg = MsgBox("*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Information source: DHS-7823 Form - Authorization to Obtain Financial Information from the Account Validation Service (AVS)." & vbcr & vbcr & _
            "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _ 
            "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _ 
            "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country.", vbInformation, "Tips and Tricks")        
            err_msg = "LOOP" & err_msg
        End if 
        If ButtonPressed = help_button_2 then 
            tips_tricks_msg = MsgBox("***" & dialog_text & "Status Date***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & help_button_2_text, vbInformation, "Tips and Tricks")        
            err_msg = "LOOP" & err_msg
        End if 
        
        'mandatory fields for all AVS_membs              
        FOR item = 0 to UBound(avs_members_array, 2)
            If avs_members_array(applicant_type_const, item) = "Select one... " then err_msg = err_msg & vbcr & "*Enter the Applicant Type."
            If avs_members_array(forms_status_const, item) = "Select one... " then err_msg = err_msg & vbcr & "* Enter the " & dialog_text & "."
            If trim(avs_members_array(avs_date_const, item)) = "" or isdate(avs_members_array(avs_date_const, item)) = False then err_msg = err_msg & vbcr & "* Enter the " & dialog_text & " Status Date." 
            If avs_members_array(forms_status_const, item) = "Received - Incomplete" and trim(avs_members_array(additional_info_const)) = "" then err_msg = err_msg & vbcr & "* Enter the reason the AVS form is incomplete."
        NEXT
            
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect    
    LOOP UNTIL err_msg = ""			
    Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
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

'----------------------------------------------------------------------------------------------------TIKL time
'set_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If set_TIKL = True then Call create_TIKL("DHS-7823 - AVS Auth Form(s) have been requested for this case. Review ECF and case notes and take applicable actions.", 10, date, False, TIKL_note_text)
        
'----------------------------------------------------------------------------------------------------The case note
'Information for the case note
memb_info = " for Memb " & member_numbers ' for the case note
If initial_option = "AVS Forms" then case_note_header = "--AVS Forms for " & HC_process & "Information--"
If initial_option = "AVS Submission/Results" then case_note_header = "--AVS Submission/Results for " & HC_process & "--"

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(case_note_header) 
If initial_option = "AVS Forms" then
    Call write_variable_in_CASE_NOTE("---Required AVS Forms Member Information---")
    Call write_variable_in_CASE_NOTE("-----------------------------------------------------------------------------------------------")
    'HH member array output 
    For item = 0 to ubound(avs_members_array, 2)
        If avs_members_array(checked_const, item) = 1 then 
            Call write_bullet_and_variable_in_CASE_NOTE("Name", avs_members_array(member_info_const, item))
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
            Call write_bullet_and_variable_in_CASE_NOTE("AVS Forms " & forms_text & " Date:", avs_members_array(avs_date_const, item))
            If trim(avs_members_array(additional_info_const, item)) <> "" then Call write_bullet_and_variable_in_CASE_NOTE("Additional Information:", avs_members_array(additional_info_const, item))
            Call write_variable_in_CASE_NOTE("-----") 
        End if 
    Next 
End if 
If verif_request = True then Call write_variable_in_CASE_NOTE("* Verification request sent to via ECF.")
If set_TIKL = true then Call write_variable_in_case_note(TIKL_note_text)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes) 
Call write_variable_in_CASE_NOTE("---") 
Call write_variable_in_CASE_NOTE(worker_signature) 


'Providing the option to run the avs option 
'If avs_option = "Received - Complete" then 
'    confirm_msgbox = msgbox("Do you wish to case note submitting the AVS request?", vbQuestion + vbYesNo, "Submit the AVS Request?")
'    If confirm_msgbox = vbYes then initial_option = "AVS Submission/Results"
'End if 

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
    script_end_procedure_with_error_report("Remember to send your verification request in ECF with the correct verbiage." & vbcr& vbcr & "This verbiage has been outputted to a Word document for your convenience.")
Else 
    script_end_procedure_with_error_report("Success! Your AVS case note has been created. Please review for accuracy & any additional information/edits.")
End if 

'    If set_TIKL = False then 
'        TIKL_msgbox = msgbox("Do you wish to send another verification requesting the AVS form?" & vbcr & vbcr & "Selecting YES will create another 10-day TIKL and case note that a verification request is being sent.", vbQuestion + vbYesNo, "Set another TIKL?")
'        If TIKL_msgbox = vbYes then 
'            set_TIKL = True
'            verif_request = True 
'        End if     
'        If TIKL_msgbox = vbNo then set_TIKL = False
'    End if 
'

'Text 20, 40, 110, 10, "Test Person Number One"
'DropListBox 140, 40, 60, 10, "Select one..."+chr(9)+"Applying"+chr(9)+"Deeming "+chr(9)+"Not Applying"+chr(9)+"Spouse", drop_one
'DropListBox 215, 40, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", List8
'EditBox 325, 40, 50, 15, date_one
'ButtonGroup ButtonPressed
'PushButton 380, 40, 15, 15, "CD", CD_one
'Text 20, 60, 110, 10, "Test Person Number Two"
'DropListBox 140, 60, 60, 10, "Select one..."+chr(9)+"Applying"+chr(9)+"Deeming "+chr(9)+"Not Applying"+chr(9)+"Spouse", Drop_two
'DropListBox 215, 60, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", forms_two
'EditBox 325, 60, 50, 15, date_two
'ButtonGroup ButtonPressed
'PushButton 380, 60, 15, 15, "CD", CD_two
'Text 20, 80, 110, 10, "Test Person Number Three"
'DropListBox 140, 80, 60, 10, "Select one..."+chr(9)+"Applying"+chr(9)+"Deeming "+chr(9)+"Not Applying"+chr(9)+"Spouse", drop_three
'DropListBox 215, 80, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", forms_three
'EditBox 325, 80, 50, 15, date_three
'ButtonGroup ButtonPressed
'PushButton 380, 80, 15, 15, "CD", CD_three
'Text 20, 100, 110, 10, "Test Person Number Four"
'DropListBox 140, 100, 60, 10, "Select one..."+chr(9)+"Applying"+chr(9)+"Deeming "+chr(9)+"Not Applying"+chr(9)+"Spouse", Drop_four
'DropListBox 215, 100, 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", forms_four
'EditBox 325, 100, 50, 15, date_four
'ButtonGroup ButtonPressed
'PushButton 380, 100, 15, 15, "CD", CD_four
