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

initial_option = "AVS Forms" 'Testing code

'----------------------------------------------------------------------------------------------------Initial dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 151, 75, "AVS Initial Process Dialog"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  DropListBox 60, 35, 85, 15, "Select one..."+chr(9)+"AVS Forms"+chr(9)+"AVS Submission", initial_option
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

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privilged, and you do not have access. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member information 
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
ReDim all_clients_array(total_clients, 2)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	all_clients_array(x, 1) = 1    '1 = checked
NEXT

If initial_option = "AVS Forms" then selection_text = "Select all members REQUIRED to sign AVS form(s):"
If initial_option = "AVS Submission" then selection_text = "Select all members AVS is being submitted for:"

Dialog1 = ""
BEGINDIALOG Dialog1, 0, 0, 241, (50 + (total_clients * 15)), "Member Selection Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
    Text 5, 5, 180, 10, selection_text
    PushButton 170, 0, 10, 15, "!", help_button
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
    Text 5, (10 + (i * 15)), 70, 10, "Select process status:"
    If initial_option = "AVS Forms" then DropListBox 80, (5 + (i * 15)), 90, 15, "Select one..."+chr(9)+"Initial Request"+chr(9)+"Not Received"+chr(9)+"Received - Complete"+chr(9)+"Received - Incomplete", avs_option 
    If initial_option = "AVS Submission" then DropListBox 80, (5 + (i * 15)), 90, 15, "Select one..."+chr(9)+"Submitting a Request"+chr(9)+"Review Results"+chr(9)+"Results After Decision", avs_option 
	ButtonGroup ButtonPressed
    PushButton 250, 170, 10, 15, "!", help_button
	OkButton 185, 10, 50, 15
	CancelButton 185, 30, 50, 15
ENDDIALOG

'TODO: add handling for no members checked 
'TODO: Update HH comp to include memb 01 if 18 or older, spouses, AREP, no SSN's

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
        If avs_option = "Select one..." then err_msg = err_msg & vbcr & "* Select an AVS process option."
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect    
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
CALL Navigate_to_MAXIS_screen("STAT", "TYPE")   'navigating to stat type to get HC application status     
    
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
        ''adding TYPE Information to be output into the dialog and case note 
        'row = 6
        'Do
        '    EmReadscreen type_memb_number, 2, row, 3
        '    EmReadscreen hc_type, 1, row, 37
        '    If hc_type = "_" then hc_type = "N"
        '    If type_memb_number = left(all_clients_array(i, 0), 2) then
        '        msgbox "match"
        '        avs_members_array(hc_type_const, avs_membs) = hc_type
        '        exit do
        '    Else 
        '        row = row + 1
        '    End if 
        'Loop until trim(type_memb_number) = ""
	END IF
NEXT

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
    BeginDialog Dialog1, 0, 0, 271, (150 + (avs_membs * 15)), "AVS Forms for #" & MAXIS_case_number
    Text 15, 10, 55, 10, date_text & " date:"
    EditBox 60, 5, 55, 15, sent_recd_date
    Text 145, 10, 45, 10, "MA Process:"
    DropListBox 195, 5, 70, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change in Basis", ma_process
    'DropListBox 195, 5, 70, 15, "Select one..."+chr(9)+"Application"+chr(9)+"Change in Basis"+chr(9)+"Renewal", ma_process
    Text 5, 30, 50, 10, "Request Type:"
    
    If avs_option = "Received - Incomplete" then
        Text 5, 50, 65, 10, "Reason Incomplete:"
        EditBox 75, 45, 190, 15, reason_incomplete 
    End if 
      x = 0
      FOR item = 0 to ubound(avs_members_array, 2)							'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
          If avs_members_array(member_name_const, item) <> "" then Text 10, (80 + (x * 15)), 140, 10, " - " & avs_members_array(member_name_const, item)
          If avs_members_array(hc_type_const, item) <> "" then DropListBox 60, 25, 150, 15, "Select One:"+chr(9)+"BI-Brain Injury Waiver"+chr(9)+"BX-Blind"+chr(9)+"CA-Community Alt. Care"+chr(9)+"DD-Developmental Disa Waiver"+chr(9)+"DP-MA for Employed Pers w/ Disa"+chr(9)+"DX-Disability"+chr(9)+"EH-Emergency Medical Assistance"+chr(9)+"EW-Elderly Waiver"+chr(9)+"EX-65 and Older"+chr(9)+"LC-Long Term Care"+chr(9)+"MP-QMB SLMB Only"+chr(9)+"QI-QI"+chr(9)+"QW-QWD"+chr(9)+"Not Applying", request_type
          'If avs_members_array(member_name_const, item) <> "" then DropListBox 160, (110 + (x * 15)), 50, 15, "Select one..."+chr(9)+"MA"+chr(9)+"MCRE"+chr(9)+"IA"+chr(9)+"QHP", avs_members_array(hc_type_const, item)
          x = x + 1
      NEXT
      GroupBox 5, 30, 260, (20 + (x * 12)), "AVS DHS-7823 Signatures are Required for the following members:"
      Text 5, (100 + (x * 12)), 40, 10, "Other notes:"
      EditBox 50, (95 + (x * 12)), 215, 15, other_notes
      Text 5,  (120 + (x * 12)), 60, 10, "Worker Signature:" 
      EditBox 70, (115 + (x * 12)), 85, 15, worker_signature
      ButtonGroup ButtonPressed
      OkButton 160, (115 + (x * 12)), 50, 15
      CancelButton 215, (115 + (x * 12)), 50, 15
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
'
'    '(initial_option = "2. Request to end eligibility in METS" or "3. Eligibility ended in METS")
'    Dialog1 = ""
'    BeginDialog Dialog1, 0, 0, 271, (120 + (avs_membs * 15)), "Eligibility Ending for #" & MAXIS_case_number
'    If initial_option = "3. Eligibility ended in METS" then 
'        Text 10, 10, 70, 10, "MMIS elig end date:"
'        EditBox 80, 5, 55, 15, mmis_end_date
'    End if 
'      Text 140, 10, 70, 10, "METS Case Number:"
'      EditBox 210, 5, 55, 15, METS_case_number
'      x = 0
'      FOR item = 0 to ubound(avs_members_array, 2)							'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
'          Text 10, (40 + (x * 15)), 100, 10, avs_members_array(member_name_const, item)
'          x = x + 1
'      NEXT
'      GroupBox 5, 25, 260, (25 + (x * 10)), "Client(s) name"
'      Text 5, (65 + (x * 10)), 40, 10, "Other notes:"
'      EditBox 50, (60 + (x * 10)), 215, 15, other_notes
'      Text 5, (85 + (x * 10)), 60, 10, "Worker Signature:"
'      EditBox 70, (80 + (x * 10)), 85, 15, worker_signature
'      ButtonGroup ButtonPressed
'          OkButton 160, (80 + (x * 10)), 50, 15
'          CancelButton 215, (80 + (x * 10)), 50, 15
'    EndDialog
'
'    DO
'        DO
'            err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
'            dialog Dialog1				    
'            cancel_without_confirmation              'new function that will cancel, collect stats, but not give user option to confirm ending script.
'            If (initial_option = "3. Eligibility ended in METS" AND (isdate(mmis_end_date) = false or trim(mmis_end_date) = "")) then err_msg = err_msg & vbcr & "* Enter a valid MMIS end date."
'            If trim(METS_case_number) = "" or IsNumeric(METS_case_number) = False or len(METS_case_number) <> 8 then err_msg = err_msg & vbcr & "* Enter a valid METS case number."
'            IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
'            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
'        LOOP UNTIL err_msg = ""									'loops until all errors are resolved
'        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
'    Loop until are_we_passworded_out = false					'loops until user passwords back in
'End if 

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

'Providing the option to run the avd option 
'If avs_option = "Received - Complete" then confirm_msgbox = msgbox("Do you wish to case note submitting the AVS request?", vbQuestion + vbYesNo, "Submit the AVS Request?")
'If confirm_msgbox = vbYes then 
'If confirm_msgbox = vbNo then 
    

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