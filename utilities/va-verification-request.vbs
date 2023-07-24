'GATHERING STATS===========================================================================================
name_of_script = "UTILITIES - VA VERIFICATION REQUEST.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
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
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("11/25/2020", "Inital Version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "VA verification email has been sent." 'setting up closing_message variable for possible additions later based on conditions

'---------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 111, 45, "Case Number"
  EditBox 65, 5, 40, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 20, 25, 40, 15
    CancelButton 65, 25, 40, 15
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------Gathering the member
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = TRUE then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

client_array = "Select One:" & "|"
DO		'reads the reference number, last name, first name, and THEN puts it into a single string THEN into the array
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
    EMReadscreen MEMB_number, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EMReadscreen client_SSN, 11, 7, 42
    client_SSN = replace(client_SSN, " ", "")
    last_name = trim(replace(last_name, "_", "")) & " "
    first_name = trim(replace(first_name, "_", "")) & " "
    mid_initial = replace(mid_initial, "_", "")
    client_string = first_name & last_name & client_SSN
    client_array = client_array & trim(client_string) & "|"
    TRANSMIT
    Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on he bottom row.
client_array = TRIM(client_array)
client_selection = split(client_array, "|")
CALL convert_array_to_droplist_items(client_selection, hh_member_dropdown)

DO   'loop for the HH member does need to re-read jsut need to allow us to chose'
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 171, 60, "HH Composition"
    DropListBox 5, 20, 160, 15, hh_member_dropdown, vet_member
      ButtonGroup ButtonPressed
        OkButton 70, 40, 45, 15
        CancelButton 120, 40, 45, 15
      Text 5, 5, 165, 10, "Please select the HH Member:"
    EndDialog

    DO
        DO
           	err_msg = ""
           	Dialog Dialog1
           	cancel_without_confirmation
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
           LOOP UNTIL err_msg = ""
    	CALL check_for_password_without_transmit(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = false

    vet_ssn = right(vet_member, 9)

    vet_member = trim(vet_member)
    vet_member = left(vet_member, len(vet_member) - 9)

    BeginDialog Dialog1, 0, 0, 226, 185, "VETERANS BENEFITS: " & maxis_case_number
      Text 10, 15, 205, 10, "Name:  " & vet_member
      Text 10, 35, 205, 10, "SSN of Veteran: " & vet_SSN
      EditBox 80, 50, 45, 15, VA_file_number
      EditBox 10, 80, 165, 15, spouse_child_name
      EditBox 10, 110, 70, 15, spouse_child_ssn
      EditBox 10, 140, 70, 15, relationship_veteran
      ButtonGroup ButtonPressed
        OkButton 135, 165, 40, 15
        CancelButton 180, 165, 40, 15
      CheckBox 10, 165, 115, 10, "Multiple requests on this case?", MULTIPLE_CHECKBOX
      GroupBox 5, 5, 215, 155, "Veteran Information:"
      Text 10, 55, 65, 10, "VA File # (if known):"
      Text 10, 70, 190, 10, "Name of Spouse/Child receiving VA benefit (if applicable):"
      Text 10, 130, 130, 10, "Relationship to Veteran (if applicable): "
      Text 10, 100, 190, 10, "SSN of Spouse/Child receiving VA benefit (if applicable): "
      ButtonGroup ButtonPressed
        PushButton 140, 50, 70, 15, "Tips and Tricks", help_button
    EndDialog

    Do
        dialog dialog1
        cancel_confirmation
        If ButtonPressed = help_button then
            tips_tricks_msg = MsgBox("*** VA File Number ***" & vbNewLine & "--------------------------" & vbNewLine & vbNewLine & "The number will appear n VA correspondence usually in the upper right corner. If a surviving spouse an XC will sometimes appear in the beginning of the file number. "  & vbNewLine & "Also, if its an older claim you may see a C in the beginning of the file number.", vbInformation, "Tips and Tricks")
        End if
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    'Creating the email
    VA_info = "Name of Veteran:  " & vet_member & vbcr & "SSN of Veteran: " & client_SSN & vbcr

    If trim(VA_file_number) <> "" THEN VA_info = VA_info & "VA File # (if known): " & VA_file_number & vbcr
    If trim(spouse_child_name) <> "" THEN VA_info = VA_info & "Name of Spouse/Child receiving VA benefit(if applicable): " & trim(spouse_child_name) & vbcr
    If trim(spouse_child_ssn) <> "" THEN VA_info = VA_info & "SSN of Spouse/Child receiving VA benefit(if applicable): " & trim(spouse_child_ssn) & vbcr
    If trim(relationship_veteran) <> "" THEN VA_info = VA_info & "Relationship to Veteran(if applicable): " & relationship_veteran

    'Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachmentsend_email)
    Call create_outlook_email("Vetservices@Hennepin.us", "", "VA Request for Case #" & MAXIS_case_number, VA_info, "", TRUE)   'will create email, will not send.

    IF MULTIPLE_CHECKBOX = 0 THEN EXIT DO 'need to be done if we are not requesting more than one '

    erase client_selection
    first_name = ""
    last_name = ""
    vet_SSN_ssn = ""
    VA_file_number = ""
    spouse_child_name = ""
    spouse_child_ssn = ""
    relationship_veteran = ""
    MULTIPLE_CHECKBOX = 0
Loop

script_end_procedure_with_error_report(closing_message)
