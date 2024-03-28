'STATS GATHERING=============================================================================================================
name_of_script = "ADMIN - QC RESULTS.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
start_time = timer
STATS_counter = 1
STATS_manualtime = 0
STATS_denominatinon = "C"

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
call changelog_update("03/27/2024", "Improve handling to determine if WCOM has been created already.", "Mark Riegel, Hennepin County")
call changelog_update("06/21/2019", "Added program selection to initial dialog. Added search for MFIP in notices for WCOM only option.", "Ilse Ferris, Hennepin County")
call changelog_update("06/21/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Name 	     Phone
'Abdi Ugas	 651-431-4026
'Chi Yang	 651-431-3964
'Erin Good	 651-431-3984
'Gary Lesney 651-431-3983
'Khin Win	 651-431-5609
'Lisa Enstad 651-431-4115
'Lor Yang	 651-431-6304
'Lori Bona	 651-431-3950
'Yer Yang	 651-431-3965

'THE SCRIPT==================================================================================================================
EMConnect ""    'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabs the MAXIS case number automatically
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 236, 70, "Enter the initial case information"
  EditBox 70, 5, 45, 15, MAXIS_case_number
  DropListBox 175, 5, 55, 15, "Select One..."+chr(9)+"MF - FS"+chr(9)+"SNAP "+chr(9)+"UHFS", program_droplist
  EditBox 70, 25, 20, 15, MAXIS_footer_month
  EditBox 95, 25, 20, 15, MAXIS_footer_year
  DropListBox 175, 25, 55, 15, "Select One..."+chr(9)+"CAPER"+chr(9)+"QC"+chr(9)+"QC < $38"+chr(9)+"WCOM only", error_selection
  EditBox 70, 50, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 150, 50, 40, 15
    CancelButton 190, 50, 40, 15
  Text 130, 30, 40, 10, "Error Type:"
  Text 20, 10, 50, 10, "Case Number: "
  Text 5, 55, 60, 10, "Worker Signature:"
  Text 5, 30, 65, 10, "Footer month/year:"
  Text 135, 10, 35, 10, "Program:"
EndDialog
'the dialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If program_droplist = "Select One..." then err_msg = err_msg & vbNewLine & "* Select the program."
        If error_selection = "Select One..." then err_msg = err_msg & vbNewLine & "* Select the error type."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If error_selection = "WCOM only" then
    STATS_manualtime = 90
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 186, 50, "Select the DHS QC Contact"
      ButtonGroup ButtonPressed
        OkButton 100, 30, 40, 15
        CancelButton 140, 30, 40, 15
      DropListBox 100, 10, 80, 15, "Select one..."+chr(9)+"Abdi Ugas"+chr(9)+"Chi Yang"+chr(9)+"Erin Good"+chr(9)+"Gary Lesney"+chr(9)+"Khin Win"+chr(9)+"Lisa Enstad"+chr(9)+"Lor Yang"+chr(9)+"Lori Bona"+chr(9)+"Yer Yang", QC_contact
      Text 5, 15, 90, 10, "Select the DHS QC contact:"
    EndDialog
    Do
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_without_confirmation
            If QC_contact = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the applicable QC contact."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    If QC_contact = "Abdi Ugas"   then phone_number = "651-431-4026"
    If QC_contact = "Chi Yang"    then phone_number = "651-431-3964"
    If QC_contact = "Erin Good"   then phone_number = "651-431-3984"
    If QC_contact = "Gary Lesney" then phone_number = "651-431-3983"
    If QC_contact = "Khin Win"    then phone_number = "651-431-5609"
    If QC_contact = "Lisa Enstad" then phone_number = "651-431-4115"
    If QC_contact = "Lor Yang"    then phone_number = "651-431-6304"
    If QC_contact = "Lori Bona"   then phone_number = "651-431-3950"
    If QC_contact = "Yer Yang"    then phone_number = "651-431-3965"

    'Navigating to the spec wcom screen
    MAXIS_background_check
    CALL Check_for_MAXIS(false)
    CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

    Emwritescreen MAXIS_footer_month, 19, 54
    Emwritescreen MAXIS_footer_year, 19, 57
    transmit

    'prog_type based on program selected in initial dialog
    If program_droplist = "MF - FS" then
        prog_type = "MF"
    else
        prog_type = "FS"
    End if

    'Searching for waiting SNAP notice
    wcom_row = 6
    Do
    	wcom_row = wcom_row + 1
    	Emreadscreen program_type, 2, wcom_row, 26
    	Emreadscreen print_status, 7, wcom_row, 71
    	If program_type = prog_type then
    		If print_status = "Waiting" then
    			Emwritescreen "x", wcom_row, 13
    			Transmit
    			PF9
    			Emreadscreen fs_wcom_exists, 3, 3, 17
    			If fs_wcom_exists <> "   " then script_end_procedure ("It appears you already have a WCOM added to this notice. The script will now end.")
    			fs_wcom_writen = true
    			'Writing in the SPEC/WCOM verbiage
    			CALL write_variable_in_SPEC_MEMO("******************************************************")
    			CALL write_variable_in_SPEC_MEMO("What to do next:")
    			CALL write_variable_in_SPEC_MEMO("You will need to contact " & QC_contact & " at the State Quality Control Office to find out what you need to do to cooperate. The phone number is " & phone_number & ".")
    			CALL write_variable_in_SPEC_MEMO("******************************************************")
    			PF4
    			PF3
    		End If
    	End If
    	If fs_wcom_writen = true then Exit Do
    	If wcom_row = 17 then
    		PF8
    		Emreadscreen spec_edit_check, 6, 24, 2
    		wcom_row = 6
    	end if
    	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
    Loop until spec_edit_check = "NOTICE"

    If no_fs_waiting = true then script_end_procedure("No waiting FS notice was found for the requested month. The script will now end.")

    script_end_procedure("WCOM has been added to the first found waiting SNAP notice for the month and case selected. Please review the notice.")
Else
    STATS_manualtime = 180
    contact_date = date & ""
    reminder_checkbox = 1 'checked

    If error_selection = "CAPER" then height = 90
    If error_selection = "QC < $38" then height = 135
    If error_selection = "QC" then height = 160

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 371, height, error_selection & " Information for case #" & MAXIS_case_number
      ButtonGroup ButtonPressed
      text 15, 15, 50, 10, "CM Reference:"
      EditBox 70, 10, 295, 15, CM_reference
      Text 5, 40, 60, 10, "HSR error sent to:"
      EditBox 70, 35, 90, 15, HSR_name
      Text 165, 40, 10, 10, "on"
      EditBox 180, 35, 50, 15, contact_date
      CheckBox 25, 55, 205, 10, "Check here add a 20 day reminder to your Outlook calendar.", reminder_checkbox
      Text 10, 75, 60, 10, "Other error notes:"
      EditBox 70, 70, 245, 15, other_notes

      If error_selection <> "CAPER" then
        Text 10, 100, 60, 10, "Error Description:"
        EditBox 70, 95, 245, 15, error_description
        Text 5, 120, 60, 10, "Agency or Client?:"
        DropListBox 70, 115, 50, 15, "Select one..."+chr(9)+"Agency"+chr(9)+"Client", agency_or_client
        Text 130, 120, 65, 10, "Payment error type:"
        DropListBox 200, 115, 65, 15, "Select one..."+chr(9)+"Overpayment"+chr(9)+"Underpayment", payment_type
      End if
      If error_selection = "QC" then
        Text 30, 145, 35, 10, "Error type:"
        EditBox 70, 140, 245, 15, error_type
      End if
      ButtonGroup ButtonPressed
        OkButton 315, 30, 50, 15
        CancelButton 315, 50, 50, 15
    EndDialog

    Do
        If error_selection = "CAPER" then
            Do
                err_msg = ""
                Dialog Dialog1
                cancel_confirmation
                If trim(CM_reference) = "" then err_msg = err_msg & vbNewLine & "* Enter the CM references."
                If trim(HSR_name) = "" then err_msg = err_msg & vbNewLine & "* Enter a HSR's name."
                If isdate(contact_date) = False or trim(contact_date) = "" then err_msg = err_msg & vbNewLine & "* Enter the date of contact with the HSR."
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
            LOOP UNTIL err_msg = ""
        ElseIf error_selection = "QC" then
            Do
                err_msg = ""
                Dialog QC_dialog
                cancel_confirmation
                If trim(CM_reference) = "" then err_msg = err_msg & vbNewLine & "* Enter the CM references."
                If trim(HSR_name) = "" then err_msg = err_msg & vbNewLine & "* Enter a HSR's name."
                If isdate(contact_date) = False or trim(contact_date) = "" then err_msg = err_msg & vbNewLine & "* Enter the date of contact with the HSR."
                If trim(error_description) = "" then err_msg = err_msg & vbNewLine & "* Enter a description of the error."
                If agency_or_client = "Select one..." then err_msg = err_msg & vbNewLine & "* Is error an agency or client error?"
                If payment_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the payment error type?"
                If trim(error_type) = "" then err_msg = err_msg & vbNewLine & "* Enter the error type from the form."
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
            LOOP UNTIL err_msg = ""
        Elseif error_selection = "QC < $38" then
            Do
                err_msg = ""
                Dialog QC_dialog
                cancel_confirmation
                If trim(CM_reference) = "" then err_msg = err_msg & vbNewLine & "* Enter the CM references."
                If trim(HSR_name) = "" then err_msg = err_msg & vbNewLine & "* Enter a HSR's name."
                If isdate(contact_date) = False or trim(contact_date) = "" then err_msg = err_msg & vbNewLine & "* Enter the date of contact with the HSR."
                If trim(error_description) = "" then err_msg = err_msg & vbNewLine & "* Enter a description of the error."
                If agency_or_client = "Select one..." then err_msg = err_msg & vbNewLine & "* Is error an agency or client error?"
                If payment_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the payment error type?"
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
            LOOP UNTIL err_msg = ""
        End if
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    reminder_date = dateadd("d", 20, contact_date)  'Setting the reminder date & Outlook appointment is created in prior to the case note
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "QC case due for " & MAXIS_case_number, "Has " & HSR_name & " returned this case?", "", TRUE, 5, "")

    Call start_a_blank_case_note
    Call write_variable_in_case_note("***State Quality Control " & error_selection & " Error for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "***")
    CALL write_bullet_and_variable_in_case_note("Progam", program_droplist)
    CALL write_bullet_and_variable_in_case_note("Combined Manual reference", CM_reference)
    CALL write_bullet_and_variable_in_case_note("Sent error to ", HSR_name)
    CALL write_bullet_and_variable_in_case_note("Due Date for resolution", reminder_date)
    Call write_variable_in_case_note("---")
    CALL write_bullet_and_variable_in_case_note("Error description", error_description)
    CALL write_bullet_and_variable_in_case_note("Agency or Client error", agency_or_client)
    CALL write_bullet_and_variable_in_case_note("Payment error type", payment_type)
    CALL write_bullet_and_variable_in_case_note("Error type", error_type)
    CALL write_bullet_and_variable_in_case_note("Other error notes", other_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)

    script_end_procedure("")
End if
