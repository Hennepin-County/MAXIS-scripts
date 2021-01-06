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
call changelog_update("03/11/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

initial_option = "1. Initial Form Request"    'Testing code
MAXIS_case_number = "298531"    'Testing code 
'avs_option = "Initial Request"  'Testing code 

'----------------------------------------------------------------------------------------------------Initial dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 151, 75, "AVS Initial Process Dialog"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  DropListBox 60, 35, 85, 15, "Select one..."+chr(9)+"1. Initial Form Request"+chr(9)+"2. AVS Form Status"+chr(9)+"3. Initial AVS Submission"+chr(9)+"4. AVS Results", initial_option
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
		dialog Dialog1				    
		cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
		call validate_MAXIS_case_number(err_msg, "*")
        If initial_option = "Select one..." then err_msg = err_msg & vbcr & "* Select an AVS process option."
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
ReDim avs_members_array(5, 0)

const member_number_const   = 0
const member_info_const     = 1 
const member_name_const     = 2 
const checked_const         = 3
const hc_applicant_const    = 4
const hc_type_const         = 5

add_to_array = False    'defaulting to false 
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EmReadscreen relationship_code, 2, 10, 42
    EmReadscreen client_age, 3, 8, 76
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
     
    If relationship_code = "01" then add_to_array = True    'applicants of HC yes 
    If relationship_code = "02" then add_to_array = True    'spouses of HC applicants yes 
    
    If trim(client_age) < "18" then add_to_array = False  'under 18 are not required to sign 
    
    If add_to_array = True then   
        ReDim Preserve avs_members_array(5,     avs_membs)
        avs_members_array(member_info_const,  avs_membs) = ref_nbr & " " & last_name & first_name
        avs_members_array(member_number_const,  avs_membs) = ref_nbr
        avs_members_array(member_name_const,    avs_membs) = first_name & "" & last_name
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_members_array(hc_applicant_const,   avs_membs) = ""         'defaulted to blank until determined if applicant or not (based on STAT/TYPE)
        avs_members_array(hc_type_const,        avs_membs) = ""         'defaulted to blank until determined later in user dialog 
        avs_membs = avs_membs + 1
    End if   
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

'Handling in case no members are idenified as needing the form. Helping to reduce errors for workers. 
If avs_membs = 0 then script_end_procedure_with_error_report("No members on this case are required by policy to sign the AVS form. Please review case if necessary.")

'----------------------------------------------------------------------------------------------------AREP information if applicable to be added to array 
Call navigate_to_MAXIS_screen("STAT", "AREP")
EMReadScreen arep_name, 37, 4, 32
arep_name = replace(arep_name, "_", "")
If trim(arep_name) <> "" then  
    client_string = "AREP: " & trim(arep_name)
    client_array = client_array & trim(client_string) & "|"
    'Adding to main array 
    ReDim Preserve avs_members_array(5,     avs_membs)
    avs_members_array(member_info_const,  avs_membs) = "AREP: " & trim(arep_name)
    avs_members_array(member_number_const,  avs_membs) = ""
    avs_members_array(member_name_const,    avs_membs) = trim(arep_name)
    avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
    avs_members_array(hc_applicant_const,   avs_membs) = FALSE      'AREP's will not be applicants          
    avs_members_array(hc_type_const,        avs_membs) = "Not Applying"
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
        ReDim Preserve avs_members_array(5,     avs_membs)
        avs_members_array(member_info_const,    avs_membs) = "Sponsor: " & trim(spon_name)
        avs_members_array(member_number_const,  avs_membs) = ""
        avs_members_array(member_name_const,    avs_membs) = trim(spon_name)
        avs_members_array(checked_const,        avs_membs) = 1          'defaulted to checked
        avs_members_array(hc_applicant_const,   avs_membs) = FALSE      'AREP's will not be applicants          
        avs_members_array(hc_type_const,        avs_membs) = "Not Applying"
        avs_membs = avs_membs + 1
    End if
    transmit
    EMReadScreen last_panel, 5, 24, 2
Loop until last_panel = "ENTER"	'This means that there are no other faci panels

'msgbox "AVS Membs: " & avs_membs

Call navigate_to_MAXIS_screen("STAT", "TYPE") 
For item = 0 to Ubound(avs_members_array, 2) 
    'non_app_info = avs_members_array(member_name_const, item)
    If avs_members_array(hc_applicant_const, item) = "" then 
        'msgbox avs_members_array(member_name_const, item)
        'adding TYPE Information to be output into the dialog and case note 
        row = 6
        Do
            EmReadscreen type_memb_number, 2, row, 3
            EmReadscreen hc_type, 1, row, 37
            'msgbox  "row: " & row & vbcr & _
            '        "memb #: " & type_memb_number & vbcr & _
            '        "hc type: " & hc_type
            If type_memb_number = avs_members_array(member_number_const, item) then
                If hc_type = "Y" then
                    msgbox "match"
                    avs_members_array(hc_applicant_const, item) = True
                    avs_members_array(hc_type_const,      item) = "Applying"  
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
        msgbox "applicant: " & avs_members_array(member_name_const, item)
    End if 
Next

If initial_option = "1. Initial Form Request" then selection_text = "Verify current Health Care Applicant in MAXIS:"

'----------------------------------------------------------------------------------------------------SELECTING APPLICANTS: Based on who is coded Y on TYPE 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 185, 125, "Applicant Selection Dialog"
    Text 5, 5, 180, 10, selection_text
    ButtonGroup ButtonPressed
    'PushButton 170, 0, 10, 15, "!", help_button
    For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        If avs_members_array(hc_applicant_const, item) = True then 
            'Text 5, 25, 100, 10, avs_members_array(member_info_const, item)
            Text 10, (25 + (item * 20)), 100, 10, avs_members_array(member_info_const, item)
            'DropListBox 115, 20, 60, 15, "Select One:"+chr(9)+"Applying"+chr(9)+"Not Applying"+chr(9)+"Spouse", avs_members_array(hc_type_const, item)
            DropListBox 120, (20 + (item * 20)), 60, 15, "Select one..."+chr(9)+"Applying"+chr(9)+"Not Applying"+chr(9)+"Spouse", avs_members_array(hc_type_const, item)
        End if     
    Next 
    ButtonGroup ButtonPressed
    OkButton 85, 105, 45, 15
    CancelButton 135, 105, 45, 15
    'PushButton 250, 170, 10, 15, "!", help_button
EndDialog

Do 
    Do 
        err_msg = ""
        Dialog Dialog1      'runs the dialog that has been dynamically created. Streamlined with new functions.
        cancel_without_confirmation
        'If ButtonPressed = help_button then 
        '    tips_tricks_msg = MsgBox("*** Who Needs to Sign the Authorization Form ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine & "Information source: DHS-7823 Form - Authorization to Obtain Financial Information from the Account Validation Service (AVS)." & vbcr & vbcr & _
        '    "- People who are applying for or enrolled in MA for people who are age 65 or older, blind or have a disability," & vbNewLine & vbNewLine & _ 
        '    "- The person's spouse, unless the person is applying for or enrolled in MA-EPD, or the person has one of the following waivers: Brain Injury (BI), Community Alternative Care (CAC), Community Access for Disability Inclusion (CADI), and Developmental Disabilities (DD)." & vbNewLine & vbNewLine & _ 
        '    "- The sponsor of the person or the person's spouse. A sponsor is someone who signed an Affidavit of Support (USCIS I-864) as a condition of the person's or his or her spouse's entry to the country.", vbInformation, "Tips and Tricks")        
        '    err_msg = "LOOP" & err_msg
        'End if 
        'ensuring that users have 
        checked_count = 0
        FOR item = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
            If avs_members_array(hc_type_const, item) = "Select one..." then err_msg = err_msg & vbcr & "* Select applicant status for each member."
        NEXT
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect    
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------SELECTING WHO IS REQUIRED TO SIGN THE FORM (different than applicants)
'Text for the next dialog based on initial option 
If initial_option = "1. Initial Form Request" then selection_text = "Select all members REQUIRED to sign AVS form(s):"
'If initial_option = "AVS Submission" then selection_text = "Select all members AVS is being submitted for:"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 185, 125, "Member Selection Dialog"
    Text 5, 5, 180, 10, selection_text
    ButtonGroup ButtonPressed
    PushButton 170, 0, 10, 15, "!", help_button
    Text 5, 30, 75, 10, "Select process status:"
    If initial_option = "AVS Forms" then DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", avs_option
    If initial_option = "AVS Submission" then DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_option   
    For item = 0 to UBound(avs_members_array, 2)									'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        If avs_members_array(checked_const, item) = 1 then checkbox 10, (50 + (item * 15)), 140, 10, avs_members_array(member_info_const, item), avs_members_array(checked_const, item)
    Next 
    ButtonGroup ButtonPressed
    OkButton 85, 105, 45, 15
    CancelButton 135, 105, 45, 15
    PushButton 250, 170, 10, 15, "!", help_button
EndDialog

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
        If avs_option = "Select one..." then err_msg = err_msg & vbcr & "* Select an AVS process option."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect    
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    
    'TODO: Is this necessary anymore?
    'Revaluing the checked or selected names based on the user selection in the dialog 
    FOR item = 0 to UBound(avs_members_array, 2)										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
        If avs_members_array(checked_const, item) = 0 then avs_members_array(checked_const, item) = 0
    NEXT
    
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
CALL Navigate_to_MAXIS_screen("STAT", "TYPE")   'navigating to stat type to get HC application status 

    

'----------------------------------------------------------------------------------------------------AVS Forms processing option
If initial_option = "AVS Forms" then 
    
    If avs_option = "Initial Request" then 
        date_text = "Sent"
        verif_request = True 
        set_TIKL = True 
    Elseif avs_option = "Not Received" then 
        date_text = "Sent"
        set_TIKL = False     
        verif_request = False  
    Elseif avs_option = "Received - Complete" then 
        date_text = "Rec'd"
        set_TIKL = False
        verif_request = False
    Elseif avs_option = "Received - Incomplete" then 
        date_text = "Rec'd"
        set_TIKL = False
        verif_request = False 
    End if 
    
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 271, 170, "AVS Forms for #" & MAXIS_case_number
        Text 25, 10, 40, 10, date_text & " date:"
        EditBox 60, 5, 55, 15, sent_recd_date
        Text 150, 10, 45, 10, "MA Process:"
        DropListBox 195, 5, 70, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change in Basis", ma_process
        'DropListBox 195, 5, 70, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change in Basis"+chr(9)+"Renewal", ma_process
        GroupBox 5, 30, 260, 75, "AVS DHS-7823 Signatures are Required for the following members:"
        'Required signagure HH list
        x = 0
        FOR item = 0 to ubound(avs_members_array, 2)							'For each person/string in the first level of the array the script will create a text box for them with height dependant on their order read
            If avs_members_array(checked_const, item) = 1 then Text 10, (45 + (x * 15)), 140, 10, " - " & avs_members_array(member_info_const, item)
            If avs_members_array(hc_type_const, item) <> "" then DropListBox 60, 45, 150, 15, "Select One:"+chr(9)+"Applying"+chr(9)+"Not Applying", request_type
            'If avs_members_array(member_name_const, item) <> "" then DropListBox 160, (110 + (x * 15)), 50, 15, "Select one..."+chr(9)+"MA"+chr(9)+"MCRE"+chr(9)+"IA"+chr(9)+"QHP", avs_members_array(hc_type_const, item)
            x = x + 1
        NEXT
            'mandatory field only visable when incomplete option is selected. 
        If avs_option = "Received - Incomplete" then
            Text 5, 115, 65, 10, "Reason Incomplete:"
            EditBox 70, 110, 195, 15, reason_incomplete
        End if 
        Text 5, 135, 40, 10, "Other notes:"
        EditBox 50, 130, 215, 15, other_notes
        Text 5, 155, 60, 10, "Worker Signature:"
        EditBox 65, 150, 105, 15, worker_signature
        ButtonGroup ButtonPressed
        OkButton 175, 150, 45, 15
        CancelButton 220, 150, 45, 15
    EndDialog
 
    'Main dialog: user will input case number and member number
    DO
        DO
            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
            Dialog Dialog1				    
            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
            If isdate(sent_recd_date) = false or trim(sent_recd_date) = "" then err_msg = err_msg & vbcr & "* Enter a valid date."
            If ma_process = "Select one..." then err_msg = err_msg & vbcr & "* Enter the MA process type."
            IF request_type = "Select one..." then err_msg = err_msg & vbcr & "* Select the applicable request type."
            If (avs_process = "Received - Incomplete" AND trim(reason_incomplete) = "") then err_msg = err_msg & vbcr & "* Enter the reason the form is incomplete."
            'For item = 0 to ubound(avs_members_array, 2)	
            '	If (avs_members_array(hc_type_const, item)) = "Select one..."then err_msg = err_msg & vbCr & "* Select a health care type for each member."
            'NEXT
            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false                    'loops until user passwords back in
    
'elseif initial_option = "AVS Submission" then 


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
        
    If set_TIKL = False then 
        TIKL_msgbox = msgbox("Do you wish to send another verification requesting the AVS form?" & vbcr & vbcr & "Selecting YES will create another 10-day TIKL and case note that a verification request is being sent.", vbQuestion + vbYesNo, "Set another TIKL?")
        If TIKL_msgbox = vbYes then 
            set_TIKL = True
            verif_request = True 
        End if     
        If TIKL_msgbox = vbNo then set_TIKL = False
    End if 
    
    '----------------------------------------------------------------------------------------------------TIKL time
    'set_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    If set_TIKL = True then Call create_TIKL("DHS-7823 - AVS Auth Form(s) have been requested for this case. Review ECF and case notes and take applicable actions.", 10, date, False, TIKL_note_text)

    member_numbers = trim(member_numbers) 'trims excess spaces of member_numbers
    If right(member_numbers, 1) = "," THEN member_numbers = left(member_numbers, len(member_numbers) - 1) 'takes the last comma off of member_numbers
    
    client_name_list = trim(client_name_list) 'trims excess spaces of client_name_list
    If right(client_name_list, 1) = "," THEN client_name_list = left(client_name_list, len(client_name_list) - 1) 'takes the last comma off of client_name_list
    
    '----------------------------------------------------------------------------------------------------The case note
    'Information for the case note
    memb_info = " for Memb " & member_numbers ' for the case note
    If initial_option = "AVS Forms" then header = "-AVS Form(s) " & avs_option & " for " & ma_process & memb_info & "-"
    If initial_option = "AVS Submission" then header = ""

    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE(header) 
    Call write_bullet_and_variable_in_CASE_NOTE("Date " & date_text, sent_recd_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Request Type", request_type)
    Call write_variable_in_CASE_NOTE("---Members that are required to sign the AVS Form(s)---") 
    
    'HH member array output 
    For i = 0 to ubound(avs_members_array, 2)
        If avs_members_array(member_name_const, i) <> "" then Call write_variable_in_CASE_NOTE(" - " & avs_members_array(member_name_const, i))
    Next 
    
    Call write_bullet_and_variable_in_CASE_NOTE("Reason Form is Incomplete", reason_incomplete)
    If verif_request = True then Call write_variable_in_CASE_NOTE("* Verification request sent to via ECF.")
    If set_TIKL = true then Call write_variable_in_case_note(TIKL_note_text)
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes) 
    Call write_variable_in_CASE_NOTE("---") 
    Call write_variable_in_CASE_NOTE(worker_signature) 
End if 

''Providing the option to run the avs option 
'If avs_option = "Received - Complete" then 
'    confirm_msgbox = msgbox("Do you wish to case note submitting the AVS request?", vbQuestion + vbYesNo, "Submit the AVS Request?")
'    If confirm_msgbox = vbYes then initial_option = "AVS Submission"
'    If confirm_msgbox = vbNo then 
'        intial_option = "AVS Forms" 
'        'Exit do 
'    End if 
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
    script_end_procedure_with_error_report("")
End if 