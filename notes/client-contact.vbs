'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CLIENT CONTACT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 195          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("05/15/2020", "Removed email for follow up to DWP baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("05/12/2020", "Added phone number support into phone number field. If MAXIS Case Number can be found at the start of the script, the phone number(s) will be autofilled. Also added support in combo boxes when options are written in, and there is an error messages to resolve.", "Ilse Ferris, Hennepin County")
call changelog_update("03/23/2020", "Updated follow up support to have case noting contain more information. Please note: EGA APPROVERS follow up is ONLY to have EGA approvers review eligiblity and/or if more verifications are needed.", "Ilse Ferris, Hennepin County")
call changelog_update("03/19/2020", "Added checkboxes to support follow up calls and/or actions needed. Email automation for DWP, YET and EGA Approvers support and TIKLs for all other baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("06/19/2018", "Added TEXT OPT OUT checkbox to be used for cases that wish to opt out of receiving text messages recertification reminders.", "Ilse Ferris, Hennepin County")
call changelog_update("06/19/2018", "Added FAX to contact type, changed MNSURE IC # to METS IC #, updated look of dialog and back end mandatory field handling. Also removed message box prior to navigating to DAIL/WRIT.", "Ilse Ferris, Hennepin County")
call changelog_update("09/28/2017", "Removed call center information from bottom of dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
when_contact_was_made = date & ", " & time 'updates the "when contact was made" variable to show the current date & time]
email_team = False 'defaulting to false for most populations 

If trim(MAXIS_case_number) <> "" then
    'Gathering the phone numbers
    call navigate_to_MAXIS_screen("STAT", "ADDR")
    phone_number_list = "Select or Type|"
    EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
    If phone_number_one <> "( ___ ) ___ ____" then phone_number_list = phone_number_list & phone_number_one & "|"
    EMReadScreen phone_number_two, 16, 18, 43
    If phone_number_two <> "( ___ ) ___ ____" then phone_number_list = phone_number_list & phone_number_two & "|"
    EMReadScreen phone_number_three, 16, 19, 43
    If phone_number_three <> "( ___ ) ___ ____" then phone_number_list = phone_number_list & phone_number_three
    phone_number_array = split(phone_number_list, "|")                
    Call convert_array_to_droplist_items(phone_number_array, phone_numbers)   
End if 

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
Do
    Do
        err_msg = ""
        Do
            BeginDialog Dialog1, 0, 0, 391, 325, "Client contact"
              ComboBox 20, 65, 65, 15, "Select or Type"+chr(9)+"Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Office visit"+chr(9)+"Letter"+chr(9)+contact_type, contact_type
              DropListBox 90, 65, 45, 10, "from"+chr(9)+"to", contact_direction
              ComboBox 140, 65, 85, 15, "Select or Type"+chr(9)+"Memb 01"+chr(9)+"Memb 02"+chr(9)+"AREP"+chr(9)+"SWKR"+chr(9)+who_contacted, who_contacted
              EditBox 245, 65, 135, 15, regarding
              ComboBox 75, 85, 75, 15, phone_numbers+chr(9)+phone_number, phone_number
              EditBox 245, 85, 135, 15, when_contact_was_made
              EditBox 75, 105, 65, 15, MAXIS_case_number
              CheckBox 165, 110, 65, 10, "Used Interpreter", used_interpreter_checkbox
              Text 255, 110, 60, 10, "METS IC number:"
              EditBox 315, 105, 65, 15, METS_IC_number
              EditBox 75, 125, 305, 15, contact_reason
              EditBox 70, 155, 310, 15, actions_taken
              EditBox 60, 195, 320, 15, verifs_needed
              EditBox 60, 215, 320, 15, case_status
              EditBox 60, 235, 320, 15, other_notes
              CheckBox 5, 260, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
              CheckBox 5, 275, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
              CheckBox 5, 290, 255, 10, "TEXT OPT OUT: Client wishes to opt out renewal text message notifications.", Opt_out_checkbox
              CheckBox 260, 260, 95, 10, "Forms were sent to AREP.", Sent_arep_checkbox
              CheckBox 260, 275, 125, 10, "Needs follow up call or processing.", follow_up_needed_checkbox
              CheckBox 260, 290, 120, 10, "Send email to EGA Approvers.", email_checkbox
              EditBox 70, 305, 205, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 280, 305, 50, 15
                CancelButton 335, 305, 50, 15
                PushButton 10, 20, 25, 10, "ADDR", ADDR_button
                PushButton 35, 20, 25, 10, "AREP", AREP_button
                PushButton 60, 20, 25, 10, "MEMB", MEMB_button
                PushButton 85, 20, 25, 10, "REVW", REVW_button
                PushButton 110, 20, 25, 10, "SWKR", SWKR_Button
                PushButton 150, 20, 50, 10, "CASE/CURR", CURR_button
                PushButton 200, 20, 50, 10, "CASE/NOTE", NOTE_button
                PushButton 250, 20, 50, 10, "ELIG/SUMM", ELIG_SUMM_button
                PushButton 300, 20, 40, 10, "MEMO", MEMO_button
                PushButton 340, 20, 40, 10, "WCOM", WCOM_button
              Text 170, 90, 75, 10, "Date/Time of Contact:"
              Text 20, 110, 50, 10, "Case number: "
              Text 10, 130, 65, 10, "Reason for contact:"
              Text 20, 160, 50, 10, "Actions taken: "
              GroupBox 5, 180, 380, 75, "Additional information about case (not mandatory):"
              Text 10, 200, 50, 10, "Verifs needed: "
              Text 15, 220, 45, 10, "Case status: "
              Text 15, 240, 40, 10, "Other notes:"
              Text 5, 310, 60, 10, "Worker signature:"
              GroupBox 5, 10, 135, 25, "STAT Navigation"
              GroupBox 5, 40, 380, 110, "Contact Information:"
              Text 30, 55, 40, 10, "Contact type"
              Text 100, 55, 30, 10, "From/To"
              Text 150, 55, 65, 10, "Who was contacted"
              Text 275, 55, 70, 10, "For case note header"
              Text 230, 70, 15, 10, "Re:"
              GroupBox 145, 10, 240, 25, "CASE Navigation"
              Text 15, 90, 50, 10, "Phone Number:"
            EndDialog
    
		    DIALOG Dialog1
		    cancel_confirmation
            MAXIS_dialog_navigation
        Loop until ButtonPressed = -1
        If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
        If trim(contact_type) = "" or contact_type = "Select or Type" then err_msg = err_msg & vbcr & "* Enter the contact type."
        If trim(who_contacted) = "" or who_contacted = "Select or Type" then err_msg = err_msg & vbcr & "* Enter who was contacted."
        If trim(contact_reason) = "" then err_msg = err_msg & vbcr & "* Enter the reason for contact."
        If trim(contact_type) = "Phone call" then 
            If trim(phone_number) = "" or trim(phone_number) = "Select or Type" then err_msg = err_msg & vbcr & "* Enter the phone number."
        End if 
		If trim(when_contact_was_made) = "" then err_msg = err_msg & vbcr & "* Enter the date and time of contact."
        If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking for an active MAXIS session
Call check_for_MAXIS(False)

Call navigate_to_MAXIS_screen("CASE", "NOTE")
EmReadscreen basket_number, 7, 21, 14    'Reading basket number to determine if email or TIKL needs to be created.

email_address = ""
'Determining if emails or TIKLS are creted and if emails are true, who are they going to.
If email_checkbox = 1 then 
    email_address = email_address & "HSPH.EWS.EGA.Approvers@hennepin.us;"
    follow_up_type = "Email"
End if 

'YET baskests
If follow_up_needed_checkbox = 1 then 
    if basket_number = "X127FA5" OR _
        basket_number = "X127FA6" OR _
        basket_number = "X127FA7" OR _
        basket_number = "X127FA8" OR _
        basket_number = "X127FA9" then 
        email_team = True
        email_address = email_address & "HSPH.ES.TEAM.4651@hennepin.us;"
        follow_up_type = "Email"
    else 
        email_team = False 
        send_dail = True 
        follow_up_type = "TIKL"
    End if 
End if 

'----------------------------------------------------------------------------------------------------TIKL Time! 
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If send_dail = true then Call create_TIKL("!!PHONE CONTACT FOLLOW UP REQUIRED!! " & when_contact_was_made, 0, date, False, TIKL_note_text)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE(contact_type & " " & contact_direction & " " & who_contacted & " re: " & regarding)
If Used_interpreter_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Contact was made: " & when_contact_was_made & " w/ interpreter")
Else
	CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
End if
If trim(phone_number) <> "Select or Type" then CALL write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number)
CALL write_bullet_and_variable_in_CASE_NOTE("METS/IC number", METS_IC_number)
CALL write_bullet_and_variable_in_CASE_NOTE("Reason for contact", contact_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifs Needed", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Case Status", case_status)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
'checkbox results
IF caf_1_check = checked THEN CALL write_variable_in_CASE_NOTE("* Reminded client about the importance of submitting the CAF 1.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF Opt_out_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Case has opted out of recert text message notifications.")
IF follow_up_needed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Follow up phone call or processing is required. " & follow_up_type & " created for follow up.")
If email_checkbox = 1 THEN CALL write_variable_in_CASE_NOTE("* Email sent to EGA Approvers for eligibility review.")
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
If Opt_out_checkbox = checked then Call create_outlook_email("xlab@maxwell.syr.edu","","Renewal text message opt out for case #" & MAXIS_case_number,"","",true)

'----------------------------------------------------------------------------------------------------Email portion
If email_checkbox = 1 or email_team = True then
    'Adding case detail to the body of the email 
    email_info = contact_type & " " & contact_direction & " " & who_contacted & " re: " & regarding & vbcr & vbcr & "* Reason for Contact: " & contact_reason & vbcr
    If trim(phone_number) <> "" then email_info = email_info & "* Phone Number: " & phone_number & vbcr
    If trim(other_notes) <> "" then email_info = email_info & "* Other Notes: " & other_notes & vbcr
    If Used_interpreter_checkbox = checked then email_info = email_info & "* An interpreter was used during this phone call." & vbcr
    If send_dail = True then email_info = email_info & vbcr & "* A TIKL has been created for team follow up."
    'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
    Call create_outlook_email(email_address, "" ,"Phone call follow up and/or processing required for case #" & MAXIS_case_number, email_info, "", True)
End if 

IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")      'Navigating to TIKL only 

end_msg = ""
'If case requires followup, it will create a MsgBox (via script_end_procedure) explaining that followup is needed. This MsgBox gets inserted into the statistics database for counties using that function. This will allow counties to "pull statistics" on follow-up, including case numbers, which can be used to track outcomes.
If follow_up_needed_checkbox = checked then end_msg = end_msg & "Success! Follow-up is needed for case number: " & MAXIS_case_number & vbcr
If Opt_out_checkbox = checked then end_msg = end_msg & "The case has been updated to OPT OUT of recert text notifications. #" & MAXIS_case_number & vbcr
If email_checkbox = checked then end_msg = end_msg & "An email has been sent to the EGA approvers with details of the contact for follow-up." & vbcr
If email_team = true then end_msg = end_msg & "An email has been sent to: " & email_address & " with details of the contact for follow-up." & vbcr
If send_dail = True then end_msg = end_msg & "A TIKL has been created for team follow up."

script_end_procedure(end_msg)