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
call changelog_update("03/30/2021", "Suggested Q-Flow population basket information updated.", "Ilse Ferris, Hennepin County")
call changelog_update("06/03/2020", "Removed TIKL and email functionality for follow ups. Q-Flow ticket number field and suggested Q-Flow population information added.", "Ilse Ferris, Hennepin County")
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

If trim(MAXIS_case_number) <> "" then
    'Gathering the phone numbers
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

    phone_number_list = "Select or Type|"
    If phone_number_one <> "" Then phone_number_list = phone_number_list & phone_number_one & "|"
    If phone_number_two <> "" Then phone_number_list = phone_number_list & phone_number_two & "|"
    If phone_number_three <> "" Then phone_number_list = phone_number_list & phone_number_three & "|"
    phone_number_array = split(phone_number_list, "|")

    Call convert_array_to_droplist_items(phone_number_array, phone_numbers)
End if

'----------------------------------------------------------------------------------------------------Adding suggested Q-Flow Ticketing population for follow up work. needed during the COVID-19 PEACETIME STATE OF EMERGENCY
EmReadscreen basket_number, 7, 21, 21    'Reading basket number
suggested_population = ""                'Blanking this out. Will default to no suggestions if x number is not in this this.

If basket_number = "X127EF8" then suggested_population = "1800"
If basket_number = "X127EF9" then suggested_population = "1800"
If basket_number = "X127EG9" then suggested_population = "1800"
If basket_number = "X127EG0" then suggested_population = "1800"

If basket_number = "X127EH1" then suggested_population = "ADS"
If basket_number = "X127EH2" then suggested_population = "ADS"
If basket_number = "X127EH3" then suggested_population = "ADS"
If basket_number = "X127EH4" then suggested_population = "ADS"
If basket_number = "X127EH5" then suggested_population = "ADS"
If basket_number = "X127EH6" then suggested_population = "ADS"
If basket_number = "X127EH7" then suggested_population = "ADS"
If basket_number = "X127EJ4" then suggested_population = "ADS"
If basket_number = "X127EJ7" then suggested_population = "ADS"
If basket_number = "X127EJ8" then suggested_population = "ADS"
If basket_number = "X127EK1" then suggested_population = "ADS"
If basket_number = "X127EK2" then suggested_population = "ADS"
If basket_number = "X127EK3" then suggested_population = "ADS"
If basket_number = "X127EK4" then suggested_population = "ADS"
If basket_number = "X127EK5" then suggested_population = "ADS"
If basket_number = "X127EK6" then suggested_population = "ADS"
If basket_number = "X127EK7" then suggested_population = "ADS"
If basket_number = "X127EK8" then suggested_population = "ADS"
If basket_number = "X127EK9" then suggested_population = "ADS"
If basket_number = "X127EM1" then suggested_population = "ADS"
If basket_number = "X127EM8" then suggested_population = "ADS"
If basket_number = "X127EM9" then suggested_population = "ADS"
If basket_number = "X127EN6" then suggested_population = "ADS"
If basket_number = "X127EP3" then suggested_population = "ADS"
If basket_number = "X127EP4" then suggested_population = "ADS"
If basket_number = "X127EP5" then suggested_population = "ADS"
If basket_number = "X127EP9" then suggested_population = "ADS"
If basket_number = "X127F3F" then suggested_population = "ADS"  'MA-EPD ADS Basket
If basket_number = "X127FE5" then suggested_population = "ADS"
If basket_number = "X127FG3" then suggested_population = "ADS"
If basket_number = "X127FH4" then suggested_population = "ADS"
If basket_number = "X127FH5" then suggested_population = "ADS"
If basket_number = "X127FI2" then suggested_population = "ADS"
If basket_number = "X127FI7" then suggested_population = "ADS"
If basket_number = "X127F3U" then suggested_population = "ADS"
If basket_number = "X127F3V" then suggested_population = "ADS"

'Contacted Case Mgt
If basket_number = "X127FG6" then suggested_population = "ADS"           '"Kristen Kasem"
If basket_number = "X127FG7" then suggested_population = "ADS"           '"Kristen Kasem"
If basket_number = "X127EM3" then suggested_population = "ADS"           '"True L. or Gina G."
If basket_number = "X127EM4" then suggested_population = "ADS"            '"True L. or Gina G."
If basket_number = "X127EW7" then suggested_population = "ADS"            '"Kimberly Hill"
If basket_number = "X127EW8" then suggested_population = "ADS"            '"Kimberly Hill"
If basket_number = "X127FF4" then suggested_population = "ADS"            '"Alyssa Taylor"
If basket_number = "X127FF5" then suggested_population = "ADS"            '"Alyssa Taylor"

If basket_number = "X127ED8" then suggested_population = "Adults"
If basket_number = "X127EE1" then suggested_population = "Adults"
If basket_number = "X127EE2" then suggested_population = "Adults"
If basket_number = "X127EE3" then suggested_population = "Adults"
If basket_number = "X127EE4" then suggested_population = "Adults"
If basket_number = "X127EE5" then suggested_population = "Adults"
If basket_number = "X127EE6" then suggested_population = "Adults"
If basket_number = "X127EE7" then suggested_population = "Adults"
If basket_number = "X127EG4" then suggested_population = "Adults"
If basket_number = "X127EG5" then suggested_population = "Adults"
If basket_number = "X127EH8" then suggested_population = "Adults"
If basket_number = "X127EJ1" then suggested_population = "Adults"
If basket_number = "X127EL1" then suggested_population = "Adults"
If basket_number = "X127EL2" then suggested_population = "Adults"
If basket_number = "X127EL3" then suggested_population = "Adults"
If basket_number = "X127EL4" then suggested_population = "Adults"
If basket_number = "X127EL5" then suggested_population = "Adults"
If basket_number = "X127EL6" then suggested_population = "Adults"
If basket_number = "X127EL7" then suggested_population = "Adults"
If basket_number = "X127EL8" then suggested_population = "Adults"
If basket_number = "X127EL9" then suggested_population = "Adults"
If basket_number = "X127EN1" then suggested_population = "Adults"
If basket_number = "X127EN2" then suggested_population = "Adults"
If basket_number = "X127EN3" then suggested_population = "Adults"
If basket_number = "X127EN4" then suggested_population = "Adults"
If basket_number = "X127EN5" then suggested_population = "Adults"
If basket_number = "X127EN7" then suggested_population = "Adults"
If basket_number = "X127EP6" then suggested_population = "Adults"
If basket_number = "X127EP7" then suggested_population = "Adults"
If basket_number = "X127EP8" then suggested_population = "Adults"
If basket_number = "X127EQ1" then suggested_population = "Adults"
If basket_number = "X127EQ2" then suggested_population = "Adults"
If basket_number = "X127EQ3" then suggested_population = "Adults"
If basket_number = "X127EQ4" then suggested_population = "Adults"
If basket_number = "X127EQ5" then suggested_population = "Adults"
If basket_number = "X127EQ8" then suggested_population = "Adults"
If basket_number = "X127EQ9" then suggested_population = "Adults"
If basket_number = "X127EX1" then suggested_population = "Adults"
If basket_number = "X127EX2" then suggested_population = "Adults"
If basket_number = "X127EX3" then suggested_population = "Adults"
If basket_number = "X127EX7" then suggested_population = "Adults"
If basket_number = "X127EX8" then suggested_population = "Adults"
If basket_number = "X127EX9" then suggested_population = "Adults"
If basket_number = "X127F3D" then suggested_population = "Adults"
If basket_number = "X127F3P" then suggested_population = "Adults"   'MA-EPD Adults Basket

If basket_number = "X127ES1" then suggested_population = "FAD"
If basket_number = "X127ES2" then suggested_population = "FAD"
If basket_number = "X127ES3" then suggested_population = "FAD"
If basket_number = "X127ES4" then suggested_population = "FAD"
If basket_number = "X127ES5" then suggested_population = "FAD"
If basket_number = "X127ES6" then suggested_population = "FAD"
If basket_number = "X127ES7" then suggested_population = "FAD"
If basket_number = "X127ES8" then suggested_population = "FAD"
If basket_number = "X127ES9" then suggested_population = "FAD"
If basket_number = "X127ET1" then suggested_population = "FAD"
If basket_number = "X127ET2" then suggested_population = "FAD"
If basket_number = "X127ET3" then suggested_population = "FAD"
If basket_number = "X127ET4" then suggested_population = "FAD"
If basket_number = "X127ET5" then suggested_population = "FAD"
If basket_number = "X127ET6" then suggested_population = "FAD"
If basket_number = "X127ET7" then suggested_population = "FAD"
If basket_number = "X127ET8" then suggested_population = "FAD"
If basket_number = "X127ET9" then suggested_population = "FAD"
If basket_number = "X127F4E" then suggested_population = "FAD"
If basket_number = "X127F3H" then suggested_population = "FAD"
If basket_number = "X127FB7" then suggested_population = "FAD"
If basket_number = "X127EZ1" then suggested_population = "FAD"
If basket_number = "X127EZ2" then suggested_population = "FAD"
If basket_number = "X127EZ3" then suggested_population = "FAD"
If basket_number = "X127EZ4" then suggested_population = "FAD"
If basket_number = "X127EZ5" then suggested_population = "FAD"
If basket_number = "X127EZ6" then suggested_population = "FAD"
If basket_number = "X127EZ7" then suggested_population = "FAD"
If basket_number = "X127EZ8" then suggested_population = "FAD"
If basket_number = "X127F3K" then suggested_population = "FAD"  'MA-EPD FAD Basket

If basket_number = "X127EH9" then suggested_population = "LTH"
If basket_number = "X127EM2" then suggested_population = "LTH"
If basket_number = "X127FE6" then suggested_population = "LTH"

If basket_number = "X127FA5" then suggested_population = "YET"
If basket_number = "X127FA6" then suggested_population = "YET"
If basket_number = "X127FA7" then suggested_population = "YET"
If basket_number = "X127FA8" then suggested_population = "YET"
If basket_number = "X127FB1" then suggested_population = "YET"
If basket_number = "X127FA9" then suggested_population = "YET"

If suggested_population = "" then suggested_population = "No suggestions available"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
Do
    Do
        err_msg = ""
        Do
            BeginDialog Dialog1, 0, 0, 391, 345, "Client contact"
            ComboBox 20, 65, 65, 15, "Select or Type"+chr(9)+"Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Office visit"+chr(9)+"Letter"+chr(9)+contact_type, contact_type
            DropListBox 90, 65, 45, 10, "from"+chr(9)+"to", contact_direction
            ComboBox 140, 65, 85, 15, "Select or Type"+chr(9)+"Memb 01"+chr(9)+"Memb 02"+chr(9)+"AREP"+chr(9)+"SWKR"+chr(9)+who_contacted, who_contacted
            EditBox 245, 65, 135, 15, regarding
            ComboBox 75, 85, 75, 15, phone_numbers+chr(9)+phone_number, phone_number
            EditBox 245, 85, 135, 15, when_contact_was_made
            EditBox 75, 105, 65, 15, MAXIS_case_number
            CheckBox 165, 110, 65, 10, "Used Interpreter", used_interpreter_checkbox
            EditBox 315, 105, 65, 15, METS_IC_number
            EditBox 75, 125, 305, 15, contact_reason
            EditBox 70, 155, 310, 15, actions_taken
            EditBox 60, 195, 320, 15, verifs_needed
            EditBox 60, 215, 320, 15, case_status
            EditBox 60, 235, 320, 15, other_notes
            CheckBox 5, 260, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
            CheckBox 5, 275, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
            CheckBox 5, 290, 255, 10, "TEXT OPT OUT: Client wishes to opt out renewal text message notifications.", Opt_out_checkbox
            CheckBox 5, 305, 95, 10, "Forms were sent to AREP.", Sent_arep_checkbox
            CheckBox 270, 260, 125, 10, "Needs follow up/hand off.", follow_up_needed_checkbox
            EditBox 340, 275, 40, 15, ticket_number                             'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
            EditBox 70, 325, 205, 15, worker_signature
            ButtonGroup ButtonPressed
            OkButton 280, 325, 50, 15
            CancelButton 335, 325, 50, 15
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
            Text 20, 110, 50, 10, "Case number: "
            Text 10, 130, 65, 10, "Reason for contact:"
            Text 20, 160, 50, 10, "Actions taken: "
            GroupBox 5, 180, 380, 75, "Additional information about case (not mandatory):"
            Text 10, 200, 50, 10, "Verifs needed: "
            Text 15, 220, 45, 10, "Case status: "
            Text 10, 330, 60, 10, "Worker signature:"
            GroupBox 5, 10, 135, 25, "STAT Navigation"
            GroupBox 5, 40, 380, 110, "Contact Information:"
            Text 30, 55, 40, 10, "Contact type"
            Text 100, 55, 30, 10, "From/To"
            Text 150, 55, 65, 10, "Who was contacted"
            Text 275, 55, 70, 10, "For case note header"
            Text 230, 70, 15, 10, "Re:"
            GroupBox 145, 10, 240, 25, "CASE Navigation"
            Text 15, 90, 50, 10, "Phone Number:"
            Text 20, 245, 40, 10, "Other notes:"
            GroupBox 260, 295, 125, 25, "Suggested Q-Flow Population:"          'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
            Text 280, 305, 100, 10, suggested_population                        'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
            Text 285, 280, 55, 10, "Q-Flow Ticket #:"                           'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
            Text 170, 90, 75, 10, "Date/Time of Contact:"
            Text 255, 110, 60, 10, "METS IC number:"
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
        If follow_up_needed_checkbox = 1 and trim(ticket_number) = "" then err_msg = err_msg & vbcr & "* Enter the Q-Flow ticket number."
        If follow_up_needed_checkbox = 0 and trim(ticket_number) <> "" then err_msg = err_msg & vbcr & "* Check the follow up box or clear the Q-flow ticket field if follow up is not needed."
        If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking for an active MAXIS session
Call check_for_MAXIS(False)
Call navigate_to_MAXIS_screen("CASE", "NOTE")

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
IF follow_up_needed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Follow up/hand off is required. Q-Flow ticket #" & ticket_number & " created.")         'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
If Opt_out_checkbox = checked then Call create_outlook_email("xlab@maxwell.syr.edu","","Renewal text message opt out for case #" & MAXIS_case_number,"","",true)

IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")      'Navigating to TIKL only

end_msg = ""
'If case requires followup, it will create a MsgBox (via script_end_procedure) explaining that followup is needed. This MsgBox gets inserted into the statistics database for counties using that function. This will allow counties to "pull statistics" on follow-up, including case numbers, which can be used to track outcomes.
If follow_up_needed_checkbox = checked then end_msg = end_msg & "Success! Follow-up is needed for case number " & MAXIS_case_number & ". Q-Flow Ticket #: " & ticket_number & vbcr
If Opt_out_checkbox = checked then end_msg = end_msg & "The case has been updated to OPT OUT of recert text notifications. #" & MAXIS_case_number & vbcr

script_end_procedure_with_error_report(end_msg)
