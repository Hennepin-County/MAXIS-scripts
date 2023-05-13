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
Call changelog_update("11/17/2022", "Added phone # & name autofilling in 'who was contacted' drop list AREP's and/or SWKR's, & added name for MEMB 01. Removed text opt out option (retired process), updated the Q-flow verbiage from N/A to NO Q-FLOW POPULATION.", "Ilse Ferris, Hennepin County")
call changelog_update("11/17/2022", "Added button to view issuance details for a case from the main dialog. This will support providing information to the resident while talking to them. This functionality does not interrupt the script run.##~####~##Look for the button that says 'Display Benefits'.##~##", "Casey Love, Hennepin County")
call changelog_update("11/14/2022", "Added button to link to the interpreter service request. Removed 'Used Interpreter' checkbox that was inactive.", "Casey Love, Hennepin County")
call changelog_update("10/07/2022", "Removed adults and families baskets that are no longer supported through Q-Flow. Any population not supported by Q-flow will now be 'N/A'.", "Ilse Ferris, Hennepin County")
call changelog_update("09/19/2022", "Suggested Q-Flow populations updated for X127EJ4 from LTC+ to Housing Supports. Removed baskets that are no longer supported through Q-Flow.", "Ilse Ferris, Hennepin County")
call changelog_update("03/03/2022", "Suggested Q-Flow population basket information added for EGA: X127EP3.", "Ilse Ferris, Hennepin County")
call changelog_update("12/10/2021", "Suggested Q-Flow population basket information added for FAD GRH: X127EZ2.", "Ilse Ferris, Hennepin County")
call changelog_update("12/03/2021", "Suggested Q-Flow population basket information added for EGA: X127EQ2 and X127EP8.", "Ilse Ferris, Hennepin County")
call changelog_update("10/07/2021", "Updated Suggested Q-Flow population to include DWP baskets (FE7,FE8, and FE9) and changed FAD poplution to Families.", "Ilse Ferris, Hennepin County")
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
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)
when_contact_was_made = date & ", " & time 'updates the "when contact was made" variable to show the current date & time]

If trim(MAXIS_case_number) <> "" then
    'Gathering the phone numbers
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

    phone_number_list = "Select or Type|"
    If phone_number_one <> "" Then phone_number_list = phone_number_list & phone_number_one & "|"
    If phone_number_two <> "" Then phone_number_list = phone_number_list & phone_number_two & "|"
    If phone_number_three <> "" Then phone_number_list = phone_number_list & phone_number_three & "|"

    '----------------------------------------------------------------------------------------------------MEMB 01 Name Collection
    Memb_01 = "Memb 01"                                 'setting value of variable, defaulting to string. 
    Call navigate_to_MAXIS_screen("STAT", "MEMB")       'navigating to STAT/MEMB. No other handling for member selection since M 01 is the default.
    EMReadScreen memb_01_check, 2, 4, 33                'ensuring it's M 01 we're reading.
    If memb_01_check = "01" then 
        EMReadScreen first_name, 12, 6, 63
        Memb_01 = "Memb 01: " & trim(replace(first_name, "_", ""))    'trim and replace underscores of the MEMB 01's 1st name; revalue MEMB 01 variable 
    End if
    
    '----------------------------------------------------------------------------------------------------AREP Name/Phone Number Collection
    case_arep = "AREP"                                  'setting value of variable, defaulting to string. 
    Call navigate_to_MAXIS_screen("STAT", "AREP")
    EmReadscreen arep_exists, 1, 2, 73                  
    If arep_exists = "1" then
        EmReadscreen arep_name, 37, 4, 32               'If an arep panel exists read the name
        case_arep = "AREP: " & trim(replace(arep_name, "_", ""))    'trim and replace underscores of the arep's name; revalue case_arep variable 

        EmReadscreen arep_phone_one, 16, 8, 32
            If arep_phone_one = "( ___ ) ___ ____" then     'If an arep phone number is not present, then establish as ""
                arep_phone_one = ""
            ELSE
                arep_phone_one = replace(arep_phone_one, "(", "")   'If not blank update the formatting
                arep_phone_one = replace(arep_phone_one, ")", "")
                arep_phone_one = trim(arep_phone_one)
                arep_phone_one = replace(arep_phone_one, " ", "-")
                phone_number_list = phone_number_list & trim(arep_phone_one) & "|" 'add to the phone_number_list that staff can choose from
            End if

        EmReadscreen arep_phone_two, 16, 9, 32
        If arep_phone_two = "( ___ ) ___ ____" then         'If an arep phone number #2 is not present, then establish as ""
            arep_phone_two = ""
        ELSE
            arep_phone_two = replace(arep_phone_two, "(", "")   'If not blank update the formatting
            arep_phone_two = replace(arep_phone_two, ")", "")
            arep_phone_two = trim(arep_phone_two)
            arep_phone_two = replace(arep_phone_two, " ", "-")
            phone_number_list = phone_number_list & trim(arep_phone_two) & "|" 'add to the phone_number_list that staff can choose from
        End if
    End if

    '----------------------------------------------------------------------------------------------------SWKR Name/Phone Number Collection
    case_SWKR = "SWKR"                                  'setting value of variable, defaulting to string. 
    Call navigate_to_MAXIS_screen("STAT", "SWKR")
    EmReadscreen SWKR_exists, 1, 2, 73                  
    If SWKR_exists = "1" then
        EmReadscreen SWKR_name, 35, 6, 32               'If an SWKR panel exists read the name
        case_SWKR = "SWKR: " & trim(replace(SWKR_name, "_", ""))    'trim and replace underscores of the SWKR's name; revalue case_SWKR variable 

        EmReadscreen SWKR_phone, 16, 12, 32
        If SWKR_phone = "( ___ ) ___ ____" then     'If an SWKR phone number is not present, then establish as ""
            SWKR_phone = ""
        ELSE
            SWKR_phone = replace(SWKR_phone, "(", "")   'If not blank update the formatting
            SWKR_phone = replace(SWKR_phone, ")", "")
            SWKR_phone = trim(SWKR_phone)
            SWKR_phone = replace(SWKR_phone, " ", "-")
            phone_number_list = phone_number_list & trim(SWKR_phone) & "|" 'add to the phone_number_list that staff can choose from
        End if
    End if     

    phone_number_array = split(phone_number_list, "|")  'creating an array of phone numbers to choose from that are active on the case, splitting by the delimiter "|"

    Call convert_array_to_droplist_items(phone_number_array, phone_numbers) 'function to add phone_number array to a droplist - variable called phone_numbers
    Call navigate_to_MAXIS_screen("STAT", "ADDR")   'navigating back to STAT/ADDR for staff to verify resident information
End if

'----------------------------------------------------------------------------------------------------Adding suggested Q-Flow Ticketing population for follow up work. needed during the COVID-19 PEACETIME STATE OF EMERGENCY
EmReadscreen basket_number, 7, 21, 21    'Reading basket number
suggested_population = "No Q-Flow Process"                'Blanking this out. Will default to no suggestions if x number is not in this this.

If basket_number = "X127EZ2" then suggested_population = "FAD GRH"

If basket_number = "X127EG5" then suggested_population = "Housing Supports"
If basket_number = "X127FG3" then suggested_population = "Housing Supports"
If basket_number = "X127EH2" then suggested_population = "Housing Supports"
If basket_number = "X127EJ4" then suggested_population = "Housing Supports"
If basket_number = "X127EJ7" then suggested_population = "Housing Supports"
If basket_number = "X127EK5" then suggested_population = "Housing Supports"
If basket_number = "X127EM1" then suggested_population = "Housing Supports"
If basket_number = "X127EM8" then suggested_population = "Housing Supports"
If basket_number = "X127EP4" then suggested_population = "Housing Supports"

If basket_number = "X127EH9" then suggested_population = "LTH"
If basket_number = "X127EJ1" then suggested_population = "LTH"
If basket_number = "X127EM2" then suggested_population = "LTH"
If basket_number = "X127FE6" then suggested_population = "LTH"

If basket_number = "X127FA5" then suggested_population = "YET"
If basket_number = "X127FA6" then suggested_population = "YET"
If basket_number = "X127FA7" then suggested_population = "YET"
If basket_number = "X127FA8" then suggested_population = "YET"
If basket_number = "X127FB1" then suggested_population = "YET"
If basket_number = "X127FA9" then suggested_population = "YET"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
Do
    Do
        err_msg = ""
        Do
            BeginDialog Dialog1, 0, 0, 391, 345, "Client Contact"
              ButtonGroup ButtonPressed
                ComboBox 20, 65, 65, 15, "Select or Type"+chr(9)+"Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Office visit"+chr(9)+"Letter"+chr(9)+contact_type, contact_type
                DropListBox 90, 65, 45, 10, "from"+chr(9)+"to", contact_direction
                ComboBox 140, 65, 85, 15, "Select or Type"+chr(9)+Memb_01+chr(9)+"Memb 02"+chr(9)+case_arep+chr(9)+case_swkr+chr(9)+who_contacted, who_contacted
                EditBox 245, 65, 135, 15, regarding
                ComboBox 75, 85, 75, 15, phone_numbers+chr(9)+phone_number, phone_number
                EditBox 245, 85, 135, 15, when_contact_was_made
                EditBox 75, 105, 45, 15, MAXIS_case_number
                PushButton 130, 105, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
                EditBox 315, 105, 65, 15, METS_IC_number
                EditBox 75, 125, 305, 15, contact_reason
                EditBox 70, 155, 310, 15, actions_taken
                EditBox 60, 195, 320, 15, verifs_needed
                EditBox 60, 215, 320, 15, case_status
                EditBox 60, 235, 320, 15, other_notes
                CheckBox 5, 260, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
                CheckBox 5, 275, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
                CheckBox 5, 290, 95, 10, "Forms were sent to AREP.", Sent_arep_checkbox
                CheckBox 270, 260, 125, 10, "Needs follow up/hand off.", follow_up_needed_checkbox
                EditBox 340, 275, 40, 15, ticket_number                             'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
                EditBox 70, 325, 205, 15, worker_signature
                OkButton 280, 325, 50, 15
                CancelButton 335, 325, 50, 15
                PushButton 10, 15, 25, 10, "ADDR", ADDR_button
                PushButton 35, 15, 25, 10, "AREP", AREP_button
                PushButton 60, 15, 25, 10, "MEMB", MEMB_button
                PushButton 85, 15, 25, 10, "REVW", REVW_button
                PushButton 110, 15, 25, 10, "SWKR", SWKR_Button
                PushButton 150, 15, 50, 10, "CASE/CURR", CURR_button
                PushButton 200, 15, 50, 10, "CASE/NOTE", NOTE_button
                PushButton 250, 15, 50, 10, "ELIG/SUMM", ELIG_SUMM_button
                PushButton 300, 15, 40, 10, "MEMO", MEMO_button
                PushButton 340, 15, 40, 10, "WCOM", WCOM_button
                Text 20, 110, 50, 10, "Case number: "
                Text 10, 130, 65, 10, "Reason for contact:"
                Text 20, 160, 50, 10, "Actions taken: "
                GroupBox 5, 180, 380, 75, "Additional information about case (not mandatory):"
                Text 10, 200, 50, 10, "Verifs needed: "
                Text 15, 220, 45, 10, "Case status: "
                Text 10, 330, 60, 10, "Worker signature:"
                GroupBox 5, 5, 135, 25, "STAT Navigation"
                GroupBox 5, 40, 380, 110, "Contact Information:"
                Text 30, 55, 40, 10, "Contact type"
                Text 100, 55, 30, 10, "From/To"
                Text 150, 55, 65, 10, "Who was contacted"
                Text 275, 55, 70, 10, "For case note header"
                Text 230, 70, 15, 10, "Re:"
                GroupBox 145, 5, 240, 25, "CASE Navigation"
                Text 15, 90, 50, 10, "Phone Number:"
                Text 15, 240, 40, 10, "Other notes:"
                GroupBox 260, 295, 125, 25, "Suggested Q-Flow Population:"          'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
                Text 280, 305, 100, 10, suggested_population                        'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
                Text 285, 280, 55, 10, "Q-Flow Ticket #:"                           'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
                Text 170, 90, 75, 10, "Date/Time of Contact:"
                Text 255, 110, 60, 10, "METS IC number:"
                PushButton 300, 35, 85, 15, "Display Benefits", display_benefits_btn
            EndDialog

		    DIALOG Dialog1
		    cancel_confirmation
            MAXIS_dialog_navigation
            If ButtonPressed = interpreter_servicves_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://itwebpw026/content/forms/af/_internal/hhs/human_services/initial_contact_access/AF10196.html"
            If ButtonPressed = display_benefits_btn Then
                display_ben_err_msg = ""
                Call validate_MAXIS_case_number(display_ben_err_msg, "*")

                If display_ben_err_msg = "" Then
                    If months_to_go_back = "" Then months_to_go_back = 3
                    run_from_client_contact = True
                    Call gather_case_benefits_details(months_to_go_back, run_from_client_contact)
                Else
                    MsgBox "Cannot display benefits details as the Case Number is not entered." & vbCr &  display_ben_err_msg
                End If
                ButtonPressed = 100
            End If
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
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")
EmReadscreen county_code, 4, 21, 14
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE:NOTE. The script will now stop.")

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE(contact_type & " " & contact_direction & " " & who_contacted & " re: " & regarding)
If Used_interpreter_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Contact was made: " & when_contact_was_made & " w/ interpreter.")
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
IF follow_up_needed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Follow up/hand off is required. Q-Flow ticket #" & ticket_number & " created.")         'needed during the COVID-19 PEACETIME STATE OF EMERGENCY
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("DAIL", "WRIT")      'Navigating to TIKL only

end_msg = ""
'If case requires followup, it will create a MsgBox (via script_end_procedure) explaining that followup is needed. This MsgBox gets inserted into the statistics database for counties using that function. This will allow counties to "pull statistics" on follow-up, including case numbers, which can be used to track outcomes.
If follow_up_needed_checkbox = checked then end_msg = end_msg & "Success! Follow-up is needed for case number " & MAXIS_case_number & ". Q-Flow Ticket #: " & ticket_number & vbcr

script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/13/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/13/2023
'--Mandatory fields all present & Reviewed--------------------------------------05/13/2023
'--All variables in dialog match mandatory fields-------------------------------05/13/2023
'Review dialog names for content and content fit in dialog----------------------05/13/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/13/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------05/13/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/13/2023
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-05/13/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/13/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------05/13/2023--------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/13/2023
'--Out-of-County handling reviewed----------------------------------------------05/13/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/13/2023
'--BULK - review output of statistics and run time/count (if applicable)--------05/13/2023--------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/13/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/13/2023
'--Incrementors reviewed (if necessary)-----------------------------------------05/13/2023
'--Denomination reviewed -------------------------------------------------------05/13/2023
'--Script name reviewed---------------------------------------------------------05/13/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/13/2023--------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/13/2023
'--comment Code-----------------------------------------------------------------05/13/2023
'--Update Changelog for release/update------------------------------------------05/13/2023
'--Remove testing message boxes-------------------------------------------------05/13/2023
'--Remove testing code/unnecessary code-----------------------------------------05/13/2023
'--Review/update SharePoint instructions----------------------------------------05/13/2023---------------Updated with new dialog & removed Text out option
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/13/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/13/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/13/2023----------------direct committing to Master branch
'--Complete misc. documentation (if applicable)---------------------------------05/13/2023
'--Update project team/issue contact (if applicable)----------------------------05/13/2023
