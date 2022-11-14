'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - 5181.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("11/14/2022", "Enhanced script to only update SWKR/Case Manager information that is added. Previously all information was cleared before updating the SWKR/case manager info.", "Ilse Ferris, Hennepin County")
call changelog_update("07/21/2022", "Fixed bug that was clearing all ADDR information.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Removed TIKL option to identify that 5181 has been rec'd.", "Ilse Ferris, Hennepin County")
call changelog_update("01/06/2020", "Updated error message handling and password handling around the dialogs.", "Ilse Ferris, Hennepin County")
call changelog_update("03/23/2018", "Updated dialog boxes to accommodate a laptop users.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number and footer month/year
EMConnect ""
call check_for_MAXIS(False) 'Checking to see that we're in MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'Showing the case number - defining the dialog for the case number
BeginDialog Dialog1 , 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, MAXIS_case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, MAXIS_footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
	OkButton 25, 45, 50, 15
	CancelButton 85, 45, 50, 15
EndDialog

DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
If is_this_priv = True then script_end_procedure("Case is privileged. The script will now end.")

'Dialog completed by worker. Each dialog follows this process:
'  1. Show the dialog and validate that next/OK or prev is pressed
'  2. Do the validation on that page, but contain a "if ButtonPressed = prev then exit do" to skip the validation if previous is pressed
'  3. Validate again that next/OK or prev is pressed
'  4. Loop until next is pressed, which will loop back to the previous dialog.

Do
    Do
    	Do
    		Do
                err_msg = ""
    			Do
                    Dialog1 = "" 'Blanking out previous dialog detail
                    'The successive dialogs for this script need to be defined in the loop just before being called
                    BeginDialog Dialog1, 0, 0, 361, 305, "5181 Dialog 1"
                      EditBox 55, 5, 55, 15, date_5181
                      EditBox 170, 5, 55, 15, date_received
                      EditBox 280, 5, 70, 15, lead_agency
                      EditBox 235, 30, 115, 15, lead_agency_assessor
                      EditBox 65, 50, 240, 15, casemgr_ADDR_line_01
                      EditBox 65, 65, 240, 15, casemgr_ADDR_line_02
                      EditBox 35, 85, 80, 15, casemgr_city
                      EditBox 155, 85, 40, 15, casemgr_state
                      EditBox 260, 85, 45, 15, casemgr_zip_code
                      EditBox 35, 105, 25, 15, phone_area_code
                      EditBox 65, 105, 25, 15, phone_prefix
                      EditBox 95, 105, 25, 15, phone_second_four
                      EditBox 140, 105, 25, 15, phone_extension
                      EditBox 190, 105, 80, 15, fax
                      CheckBox 275, 105, 80, 15, "Update SWRK panel", update_SWKR_info_checkbox
                      CheckBox 60, 140, 115, 15, "Have script update ADDR panel", update_addr_checkbox
                      EditBox 70, 160, 140, 15, name_of_facility
                      EditBox 285, 160, 65, 15, date_of_admission
                      EditBox 70, 180, 240, 15, facility_address_line_01
                      EditBox 70, 195, 240, 15, facility_address_line_02
                      EditBox 30, 215, 80, 15, facility_city
                      EditBox 140, 215, 40, 15, facility_state
                      EditBox 230, 215, 45, 15, facility_county_code
                      EditBox 310, 215, 45, 15, facility_zip_code
                      DropListBox 170, 250, 105, 15, "Select one..."+chr(9)+"No waiver"+chr(9)+"Alternative Care"+chr(9)+"BI diversion"+chr(9)+"BI conversion"+chr(9)+"CAC diversion"+chr(9)+"CAC conversion"+chr(9)+"CADI diversion"+chr(9)+"CADI conversion"+chr(9)+"DD diversion"+chr(9)+"DD conversion"+chr(9)+"EW diversion"+chr(9)+"EW conversion", waiver_type_droplist
                      CheckBox 40, 265, 190, 10, "Essential Community Supports (DHS- 3876 is required)", essential_community_supports_check
                      ButtonGroup ButtonPressed
                        PushButton 245, 285, 55, 15, "Next", next_to_page_02_button
                        CancelButton 305, 285, 50, 15
                      Text 170, 110, 20, 10, "Fax:"
                      Text 5, 160, 60, 15, "Name of Facility:"
                      Text 5, 45, 55, 15, "Address line 1:"
                      Text 5, 85, 25, 15, "City:"
                      Text 220, 160, 65, 15, "Date of admission:"
                      Text 135, 85, 20, 15, "State:"
                      Text 5, 105, 30, 10, "Phone:"
                      Text 230, 5, 45, 15, "Lead Agency:"
                      Text 225, 85, 35, 15, "Zip code:"
                      Text 5, 30, 100, 15, "**CONTACT INFORMATION**"
                      Text 5, 65, 55, 15, "Address line 2:"
                      Text 5, 180, 60, 15, "Facility address:"
                      Text 105, 30, 130, 15, "Lead Agency Assessor/Case Manager:"
                      Text 115, 5, 55, 15, "Date Received:"
                      Text 25, 250, 140, 10, "Choose waiver type (or select 'no waiver'):"
                      Text 125, 110, 15, 10, "Ext."
                      Text 30, 235, 285, 15, "OR The client is currently requesting services/enrolled in the following waiver program:"
                      Text 5, 195, 55, 15, "Address line 2:"
                      Text 5, 140, 45, 15, "**STATUS**"
                      GroupBox 0, 20, 355, 105, ""
                      Text 5, 5, 50, 15, "Date on 5181:"
                      Text 5, 215, 20, 15, "City:"
                      Text 115, 215, 20, 15, "State:"
                      Text 280, 215, 30, 15, "Zip code:"
                      GroupBox 0, 130, 355, 150, ""
                      Text 185, 215, 45, 15, "County code:"
                      Text 185, 140, 165, 15, "**Script will default to sending the SWKR notices**"
                    EndDialog

    				Dialog Dialog1						'Displays the first dialog - defined just above.
    				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
    			Loop until ButtonPressed = next_to_page_02_button
                If isdate(date_5181) = False or trim(date_5181) = "" then err_msg = err_msg & vbcr & "* Enter a valid 5181 date."
                If isdate(date_received) = False or trim(date_received) = "" then err_msg = err_msg & vbcr & "* Enter the date the 5181 was received."
                IF trim(lead_agency) = "" then err_msg = err_msg & vbcr & "* Enter the Lead Agency Name."		'Requires the user to select a waiver
                'case manager info
                If trim(casemgr_ADDR_line_01) <> "" then
                    If trim(casemgr_city) = "" then err_msg = err_msg & vBcr & "* Update the case manager's city."
                    If trim(casemgr_state) = "" then err_msg = err_msg & vBcr & "* Update the case manager's state."
                    If trim(casemgr_zip_code) = "" then err_msg = err_msg & vBcr & "* Update the case manager's zip code."
                End if
                'phone number
                If trim(phone_area_code) <> "" or trim(phone_prefix) <> "" or trim(phone_second_four) <> "" or trim(phone_extension) <> "" then
                    If trim(phone_area_code) = "" or len(phone_area_code) <> 3 then err_msg = err_msg & vBcr & "* Enter the case's managers 3-digit area code."
                    If trim(phone_prefix) = "" or len(phone_prefix) <> 3 then err_msg = err_msg & vBcr & "* Enter the case's managers 3-digit phone number prefix code."
                    If trim(phone_second_four) = "" or len(phone_second_four) <> 4 then err_msg = err_msg & vBcr & "* Enter the case's managers 4-digit phone number line code."
                End if
                'facility info
                IF update_addr_checkbox = 1 then
                    If isdate(date_of_admission) = False then err_msg = err_msg & vBcr & "* Enter the date of admission."
                    If trim(facility_address_line_01) = "" then err_msg = err_msg & vBcr & "* Update the faci address line 1."
                    If trim(facility_city) = "" then err_msg = err_msg & vBcr & "* Update the faci city."
                    If trim(facility_state) = "" then err_msg = err_msg & vBcr & "* Update the faci state."
                    If trim(facility_county_code) = "" then err_msg = err_msg & vBcr & "* Update the faci county code."
                    If trim(facility_zip_code) = "" then err_msg = err_msg & vBcr & "* Update the faci zip code."
                End if
    			If waiver_type_droplist = "Select one..." then err_msg = err_msg & vbcr & "* Choose waiver type (or select 'no waiver')."		'Requires the user to select a waiver
                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
            Loop until err_msg = ""
    	Loop until ButtonPressed = next_to_page_02_button
        ''-------------------------------------------------------------------------------------------------DIALOG
    	Do
    		Do
    			Do
                    err_msg = ""
    				Do
                        Dialog1 = "" 'Blanking out previous dialog detailThe successive dialogs for this script need to be defined in the loop just before being called
                        BeginDialog Dialog1, 0, 0, 361, 385, "5181 Dialog 2: INITIAL REQUESTS (check all that apply):"
                        EditBox 75, 15, 45, 15, waiver_assessment_date
                        EditBox 275, 30, 45, 15, estimated_effective_date
                        EditBox 120, 50, 45, 15, estimated_monthly_waiver_costs
                        CheckBox 175, 55, 170, 15, "Does not meet waiver services LOC requirement", does_not_meet_waiver_LOC_check
                        EditBox 105, 70, 60, 15, ongoing_waiver_case_manager
                        EditBox 75, 110, 45, 15, LTCF_assessment_date
                        CheckBox 130, 115, 100, 10, "Meets MA-LOC requirement", meets_MALOC_check
                        EditBox 130, 130, 110, 15, ongoing_case_manager
                        CheckBox 10, 150, 135, 10, "Ongoing case manager not available", ongoing_case_manager_not_available_check
                        CheckBox 10, 160, 115, 10, "Does not meet LOC requirement", does_not_meet_MALTC_LOC_check
                        CheckBox 150, 150, 65, 10, "1503 requested?", requested_1503_check
                        CheckBox 150, 160, 55, 10, "1503 on file?", onfile_1503_check
                        CheckBox 10, 200, 80, 15, "Client applied for MA", client_applied_MA_check
                        EditBox 240, 210, 45, 15, Client_MA_enrollee
                        CheckBox 10, 225, 195, 15, "Completed DHS-3543 or DHS-3531 attached to DHS-5181", completed_3543_3531_check
                        EditBox 235, 240, 45, 15, completed_3543_3531_faxed
                        CheckBox 10, 255, 180, 15, "Please send DHS-3543 to client (MA enrollee)", please_send_3543_check
                        EditBox 185, 270, 150, 15, please_send_3531
                        CheckBox 10, 290, 205, 10, "Please send DHS-3340 to client - Asset Assessment needed", please_send_3340_check
                        EditBox 240, 320, 45, 15, client_no_longer_meets_LOC_efffective_date
                        DropListBox 105, 340, 60, 15, "Select one..."+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EW", from_droplist
                        DropListBox 180, 340, 60, 15, "Select one..."+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EW", to_droplist
                        EditBox 295, 340, 55, 15, waiver_program_change_effective_date
                        ButtonGroup ButtonPressed
                          PushButton 190, 365, 50, 15, "Previous", previous_to_page_01_button
                          PushButton 245, 365, 50, 15, "Next", next_to_page_03_button
                          CancelButton 300, 365, 50, 15
                        GroupBox 5, 5, 350, 85, "**WAIVERS** Assessment date determine client:"
                        GroupBox 5, 100, 350, 80, "**LTCF** Assessment determines client: "
                        GroupBox 5, 190, 350, 115, "**MEDICAL ASSISTANCE REQUESTS/APPLICATIONS**"
                        Text 10, 115, 60, 10, "Assessment date:"
                        Text 10, 35, 265, 10, "Needs waiver services and meets LOC. Anticipated effective date no sooner than:"
                        Text 170, 345, 10, 10, "to:"
                        Text 10, 20, 60, 10, "Assessment date:"
                        Text 245, 345, 50, 10, "Effective date:"
                        GroupBox 5, 310, 350, 50, "**CHANGES COMPLETED BY THE ASSESSOR**"
                        Text 10, 55, 110, 10, "Estimated monthly waiver costs:"
                        Text 10, 75, 95, 10, "Ongoing case mgr assigned:"
                        Text 10, 135, 110, 10, "Ongoing case manager assigned:"
                        Text 10, 215, 230, 10, "Client is an MA enrollee -  If assessor provided DHS-3543, enter date:"
                        Text 10, 245, 225, 10, "If completed DHS-3543 or DHS-3531 was faxed to county, enter date: "
                        Text 10, 275, 170, 10, "Please send DHS-3531 to client (Not MA enrollee) at:"
                        Text 10, 325, 225, 10, "Client no longer meets LOC - Effective date should be no sooner than:"
                        Text 5, 345, 100, 10, "Waiver program change from:"
                        EndDialog
    					Dialog Dialog1							'Displays the second dialog - defined just above.
    					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
    					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
    				Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
    				If ButtonPressed = previous_to_page_01_button THEN exit do
                    If (from_droplist = "Select one..." AND to_droplist <> "Select one...") OR (from_droplist <> "Select one..." AND to_droplist = "Select one...") then err_msg = err_msg & vbcr & "You must enter valid selections for the waiver program change 'to' and 'from'." 'Requires the user to enter a droplist item
                    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
                Loop until err_msg = ""
    		Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
    		If ButtonPressed = previous_to_page_01_button then exit do

            Do
    			Do
    			    err_msg = ""
                        '-------------------------------------------------------------------------------------------------DIALOG
                        Dialog1 = "" 'Blanking out previous dialog detail
                        'The successive dialogs for this script need to be defined in the loop just before being called
                         BeginDialog Dialog1 , 0, 0, 366, 345, "5181 Dialog 3"
                         CheckBox 10, 20, 130, 10, "Exited waiver program effective date: ", exited_waiver_program_check
                         EditBox 150, 15, 40, 15, exit_waiver_end_date
                         CheckBox 15, 40, 60, 10, "Client's choice", client_choice_check
                         CheckBox 200, 20, 115, 10, "Client deceased.  Date of death:", client_deceased_check
                         EditBox 315, 15, 40, 15, date_of_death
                         CheckBox 200, 40, 95, 10, "Client moved to LTCF on:", client_moved_to_LTCF_check
                         EditBox 315, 35, 40, 15, client_moved_to_LTCF
                         EditBox 75, 55, 235, 15, LTCF_ADDR_line_01
                         EditBox 75, 75, 235, 15, LTCF_ADDR_line_02
                         EditBox 35, 95, 55, 15, LTCF_city
                         EditBox 120, 95, 25, 15, LTCF_state
                         EditBox 195, 95, 25, 15, LTCF_county_code
                         EditBox 265, 95, 45, 15, LTCF_zip_code
                         CheckBox 15, 115, 115, 10, "Have script update ADDR panel", LTCF_update_ADDR_checkbox
                         CheckBox 15, 135, 110, 10, "Waiver program change: From", waiver_program_change_check
                         EditBox 125, 130, 45, 15, waiver_program_change_from
                         EditBox 190, 130, 45, 15, waiver_program_change_to
                         CheckBox 15, 155, 175, 10, "Client disenrolled from health plan.  Effective date: ", client_disenrolled_health_plan_check
                         EditBox 190, 150, 45, 15, client_disenrolled_from_healthplan
                         CheckBox 15, 175, 105, 10, "New address-Effective date:", new_address_check
                         EditBox 125, 170, 45, 15, new_address_effective_date
                         EditBox 80, 190, 235, 15, change_ADDR_line_1
                         EditBox 80, 210, 235, 15, change_ADDR_line_2
                         EditBox 35, 230, 60, 15, change_city
                         EditBox 125, 230, 25, 15, change_state
                         EditBox 205, 230, 25, 15, change_county_code
                         EditBox 270, 230, 45, 15, change_zip_code
                         CheckBox 15, 250, 115, 10, "Have script update ADDR panel", update_addr_new_ADDR_checkbox
                         EditBox 65, 270, 285, 15, case_action
                         EditBox 65, 290, 285, 15, other_notes
                         CheckBox 20, 310, 125, 10, "Sent 5181 back to Case Manager?", sent_5181_to_caseworker_check
                         EditBox 70, 325, 120, 15, worker_signature
                         ButtonGroup ButtonPressed
                           PushButton 195, 325, 50, 15, "Previous", previous_to_page_02_button
                           OkButton 250, 325, 50, 15
                           CancelButton 305, 325, 50, 15
                         Text 15, 75, 55, 10, "Address line 2:"
                         Text 15, 100, 20, 10, "City:"
                         Text 5, 330, 65, 10, "Worker signature:"
                         Text 95, 100, 25, 10, "State:"
                         Text 175, 135, 15, 10, "To: "
                         Text 150, 100, 45, 10, "County code:"
                         Text 230, 100, 35, 10, "Zip code:"
                         Text 15, 275, 45, 10, "Case Action:"
                         Text 15, 60, 60, 10, "Facility Address:"
                         Text 15, 195, 60, 10, "Address line 1:"
                         Text 15, 215, 55, 10, "Address line 2:"
                         Text 15, 235, 20, 10, "City:"
                         Text 100, 235, 20, 10, "State:"
                         Text 155, 235, 45, 10, "County code:"
                         Text 235, 235, 35, 10, "Zip code:"
                         Text 15, 295, 45, 10, "Other notes:"
                         GroupBox 5, 5, 355, 260, "**CHANGES** (check all that apply):"
                       EndDialog
    				Dialog Dialog1							'Displays the third dialog - defined just above.
    				cancel_confirmation					'Asks if you're sure you want to cancel, and cancels if you select that.
    				MAXIS_dialog_navigation				'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
    				IF (exited_waiver_program_check = checked AND isdate(exit_waiver_end_date) = false) THEN err_msg = err_msg & vBcr & "* Complete the field next to the exited waiver checkbox that was checked."
    				IF (client_deceased_check =  checked AND isdate(date_of_death) = false) THEN err_msg = err_msg & vBcr & "* Complete the field next to the client deceased checkbox that was checked."
    				IF (client_moved_to_LTCF_check = checked AND isdate(client_moved_to_LTCF) = False) THEN err_msg = err_msg & vBcr & "* Complete the field next to the client moved to LTCF checkbox that was checked."
                    If LTCF_update_ADDR_checkbox = 1 then
                        If trim(LTCF_ADDR_line_01) = "" then err_msg = err_msg & vBcr & "* Update the faci address line 1."
                        If trim(LTCF_city) = "" then err_msg = err_msg & vBcr & "* Update the faci city."
                        If trim(LTCF_state) = "" then err_msg = err_msg & vBcr & "* Update the faci state."
                        If trim(LTCF_county_code) = "" then err_msg = err_msg & vBcr & "* Update the faci county code."
                        If trim(LTCF_zip_code) = "" then err_msg = err_msg & vBcr & "* Update the faci zip code."
                    End if
                    IF (waiver_program_change_check = checked AND waiver_program_change_from = "" AND waiver_program_change_to = "") THEN err_msg = err_msg & vBcr & "* Complete the field next to the waiver program change checkbox that was checked."
    				IF (client_disenrolled_health_plan_check = checked AND client_disenrolled_from_healthplan = "") THEN err_msg = err_msg & vBcr & "* Complete a field next to the client disenrolled from health plan checkbox that was checked."
    				IF (new_address_check = checked AND isdate(new_address_effective_date) = False) THEN err_msg = err_msg & vBcr & "* Complete a field next to the new address effective date checkbox that was checked."
                    If update_addr_new_ADDR_checkbox = 1 then
                        If trim(change_ADDR_line_1) = "" then err_msg = err_msg & vBcr & "* Update the new address line 1."
                        If trim(change_city) = "" then err_msg = err_msg & vBcr & "* Update the new city."
                        If trim(change_state) = "" then err_msg = err_msg & vBcr & "* Update the new state."
                        If trim(change_county_code) = "" then err_msg = err_msg & vBcr & "* Update the new county code."
                        If trim(change_zip_code) = "" then err_msg = err_msg & vBcr & "* Update the new zip code."
                    End if
                    IF trim(case_action) = "" THEN err_msg = err_msg & vBcr & "* Complete case actions section."
    				IF trim(worker_signature) = "" THEN err_msg = err_msg & vBcr & "* Sign your case note."
                    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    			Loop until err_msg = ""
    		Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button
    	Loop until ButtonPressed = -1
    	CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
    Loop until case_note_confirm = TRUE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call check_for_MAXIS(False) 'Checking to see that we're in MAXIS

'Dollar bill symbol will be added to numeric variables
IF estimated_monthly_waiver_costs <> "" THEN estimated_monthly_waiver_costs = "$" & estimated_monthly_waiver_costs

'ACTIONS----------------------------------------------------------------------------------------------------
'Updates STAT MEMB with client's date of death (client_deceased_check)
IF client_deceased_check = 1 THEN  	'Goes to STAT MEMB
	'Creates a new variable with MAXIS_footer_month and MAXIS_footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year
	'Calculates the difference between the two dates (date of death and footer month)
	difference_between_dates = DateDiff("m", date_of_death, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		Call convert_date_into_MAXIS_footer_month(date_of_death, MAXIS_footer_month, MAXIS_footer_year)
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		Transmit
	END IF
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	PF9
	'Writes in DOD from the date_of_death
	Call create_MAXIS_friendly_date_with_YYYY(date_of_death, 0, 19, 42)
	transmit
	PF3
	transmit
END IF

'------ADDRESS UPDATES----------------------------------------------------------------------------------------------------
'Updates ADDR if selected on DIALOG 1 "have script update ADDR panel"
IF update_addr_checkbox = 1 THEN
	'Creates a new variable with MAXIS_footer_month and MAXIS_footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", date_of_admission, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(date_of_admission, MAXIS_footer_month, MAXIS_footer_year)
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		Transmit
	END IF

    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

	Call access_ADDR_panel("WRITE", notes_on_address, facility_address_line_01, facility_address_line_02, resi_street_full, facility_city, facility_state, facility_zip_code, facility_county_code, "OT - Other Document", addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, date_of_admission, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
END If

'Updates ADDR if selected on DIALOG 3 "have script update ADDR panel" for move to LTCF
IF LTCF_update_ADDR_checkbox = 1 THEN
		'Creates a new variable with MAXIS_footer_month and MAXIS_footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", client_moved_to_LTCF, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(client_moved_to_LTCF, MAXIS_footer_month, MAXIS_footer_year)
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		Transmit
	END IF

    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

	Call access_ADDR_panel("WRITE", notes_on_address, LTCF_ADDR_line_01, LTCF_ADDR_line_02, resi_street_full, LTCF_city, LTCF_state, LTCF_zip_code, LTCF_county_code, "OT - Other Document", addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, client_moved_to_LTCF, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
END If

'Updates ADDR if selected on DIALOG 3 "have script update ADDR panel" for new address
IF update_addr_new_ADDR_checkbox = 1 THEN
	'Creates a new variable with MAXIS_footer_month and MAXIS_footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", new_address_effective_date, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(new_address_effective_date, MAXIS_footer_month, MAXIS_footer_year)
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		Transmit
	END IF

    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

	Call access_ADDR_panel("WRITE", notes_on_address, change_ADDR_line_1, change_ADDR_line_2, resi_street_full, change_city, change_state, change_zip_code, change_county_code, "OT - Other Document", addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, new_address_effective_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
END If

'Updates SWKR panel with Name, address and phone number if checked on DIALOG 1
If update_SWKR_info_checkbox = 1 THEN

	Call navigate_to_MAXIS_screen("STAT", "SWKR")  'Go to STAT/SWKR
	'creates a new panel if one doesn't exist, and will needs new if there is not one
	EMReadScreen panel_exists_check, 1, 2, 73
	IF panel_exists_check = "0" THEN
		EMWriteScreen "NN", 20, 79 'creating new panel
		transmit
	ELSE
		PF9	'putting panel into edit mode
	END IF

	'Blanks out the old info and writes in the new info into the SWKR panel if updated in the dialog
    If trim(lead_agency_assessor) <> "" then
        call clear_line_of_text(6, 32)
        EMWriteScreen lead_agency_assessor, 6, 32
    End if

    'updating all the ADDR info together
    If trim(casemgr_ADDR_line_01) <> "" then
        call clear_line_of_text(8, 32)
        call clear_line_of_text(9, 32)
        EMWriteScreen casemgr_ADDR_line_01, 8, 32
        EMWriteScreen casemgr_ADDR_line_02, 9, 32
    End if

    If trim(casemgr_city) <> "" then
        call clear_line_of_text(10, 32)
        EMWriteScreen casemgr_city, 10, 32
    End if

    If trim(casemgr_state) <> "" then
        call clear_line_of_text(10, 54)
        EMWriteScreen casemgr_state, 10, 54
    End if

    If trim(casemgr_zip_code) <> "" then
        call clear_line_of_text(10, 63)
        EMWriteScreen casemgr_zip_code, 10, 63
    End if

    'Updating all the phone number info together
    If trim(phone_area_code) <> "" then
        call clear_line_of_text(12, 34)
        call clear_line_of_text(12, 40)
        call clear_line_of_text(12, 44)
        call clear_line_of_text(12, 54)
        EMWriteScreen phone_area_code, 12, 34
        EMWriteScreen phone_prefix, 12, 40
        EMWriteScreen phone_second_four, 12, 44
        EMWriteScreen phone_extension, 12, 54
    End if

	EMWriteScreen "Y", 15, 63
	transmit
	transmit
	PF3
END IF

'Updates SWKR panel with ongoing waiver case manager assigned
If trim(ongoing_waiver_case_manager) <> "" then
	Call navigate_to_MAXIS_screen("STAT", "SWKR")  'Go to STAT/SWKR
	PF9    'Go into edit mode
	Call clear_line_of_text(6, 32) 'Blanks out the old info
	EMWriteScreen ongoing_waiver_case_manager, 6, 32   'Writes in new case manager name
	transmit
	transmit
	PF3
END IF

'Updates SWKR panel with ongoing case manager assigned
If trim(ongoing_case_manager) <> "" then
	Call navigate_to_MAXIS_screen("STAT", "SWKR")  'Go to STAT/SWKR
	PF9    	'Go into edit mode
	Call clear_line_of_text(6, 32) 'Blanks out the old info
	EMWriteScreen ongoing_case_manager, 6, 32 'Writes in new case manager name
	transmit
	transmit
	PF3
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
'Information from DHS 5181 Dialog 1
'Contact information
Call write_variable_in_case_note ("~~~DHS 5181 rec'd~~~")
Call write_bullet_and_variable_in_case_note ("Date of 5181", date_5181 )
Call write_bullet_and_variable_in_case_note ("Date received", date_received)
Call write_bullet_and_variable_in_case_note ("Lead Agency", lead_agency)
Call write_bullet_and_variable_in_case_note ("Lead Agency Assessor/Case Manager",lead_agency_assessor)
Call write_bullet_and_variable_in_case_note ("Address", casemgr_ADDR_line_01 & " " & casemgr_ADDR_line_02 & " " & casemgr_city & " " & casemgr_state & " " & casemgr_zip_code)
Call write_bullet_and_variable_in_case_note ("Phone", phone_area_code & "-" & phone_prefix & "-" & phone_second_four & " " & phone_extension)
Call write_bullet_and_variable_in_case_note ("Fax", fax)
'STATUS
Call write_bullet_and_variable_in_case_note ("Name of Facility", name_of_facility)
Call write_bullet_and_variable_in_case_note ("Date of admission", date_of_admission)
Call write_bullet_and_variable_in_case_note ("Facility address", facility_address_line_01 & " " & facility_address_line_02 & " " & facility_city & " " & facility_state & " " & facility_zip_code)
IF waiver_type_droplist <> "No waiver" then call write_bullet_and_variable_in_case_note("Client is requesting services/enrolled in waiver type", waiver_type_droplist)
IF essential_community_supports_check = 1 THEN Call write_variable_in_case_note ("* Essential Community supports.  Client does not meet LOC requirements.")

'Information from DHS 5181 Dialog 2
'Waivers
Call write_bullet_and_variable_in_case_note ("Waiver Assessment Date", waiver_assessment_date)
Call write_bullet_and_variable_in_case_note ("Assessment determines that client needs waiver services and meets LOC requirements.  Anticipated effective date no sooner than", estimated_effective_date)
Call write_bullet_and_variable_in_case_note ("Estimated monthly waiver costs", estimated_monthly_waiver_costs)
IF does_not_meet_waiver_LOC_check = 1 THEN Call write_variable_in_case_note ("* Client does not meet LOC requirements for waivered services.")
Call write_bullet_and_variable_in_case_note ("Ongoing case manager is", ongoing_waiver_case_manager)
'LTCF
Call write_bullet_and_variable_in_case_note ("LTCF Assessment Date", LTCF_assessment_date)
IF meets_MALOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client meets the LOC requirement.")
Call write_bullet_and_variable_in_case_note("Ongoing case manager is", ongoing_case_manager)
IF ongoing_case_manager_not_available_check = 1 THEN Call write_variable_in_case_note ("* Ongoing Case Manager not available.")
IF does_not_meet_MALTC_LOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client does not meet LOC requirements for LTCF's.")
IF requested_1503_check = 1 THEN Call write_variable_in_case_note ("* A DHS-1503 has been requested from the LTCF.")
IF onfile_1503_check = 1 THEN Call write_variable_in_case_note ("A DHS-1503 has been provided.")
'MA requests/applications
IF client_applied_MA_check = 1 THEN Call write_variable_in_case_note ("* Client has applied for MA.")
Call write_bullet_and_variable_in_case_note ("Client is an MA enrollee. Assessor provided a DHS-3543 on", Client_MA_enrollee)
IF completed_3543_3531_check = 1 THEN Call write_variable_in_case_note ("* Completed DHS-3543 or DHS-3531 attached to DHS 5181.")
Call write_bullet_and_variable_in_case_note ("Completed DHS-3543 or DHS-3531 faxed to county on", completed_3543_3531_faxed)
IF please_send_3543_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested that a DHS-3543 be sent to the MA enrollee or AREP.")
Call write_bullet_and_variable_in_case_note ("* Case manager has requested that a DHS-3531 be sent to a non-MA enrollee at", please_send_3531)
IF please_send_3340_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested an Asset Assessment, DHS 3340, be send to the client or AREP.")
'changes completed by the assessor
Call write_bullet_and_variable_in_case_note ("Client no longer meets LOC - Effective date should be no sooner than", client_no_longer_meets_LOC_efffective_date)
IF from_droplist <> "Select one..." AND to_droplist <> "Select one.." THEN Call write_bullet_and_variable_in_case_note ("Waiver program changed from", from_droplist & " to: " & to_droplist & ". Effective date: " & waiver_program_change_effective_date)

'Information from DHS 5181 Dialog 3
'changes
IF exited_waiver_program_check = 1 THEN Call write_variable_in_case_note("* Exited waiver program.  Effective date: " & exit_waiver_end_date)
IF client_choice_check = 1 THEN Call write_variable_in_case_note ("* Client has chosen to exit the waiver program.")
IF client_deceased_check = 1 THEN Call write_variable_in_case_note ("* Client is deceased.  Date of death: " & date_of_death)
IF client_moved_to_LTCF_check = 1 THEN Call write_variable_in_case_note ("* Client moved to LTCF on" & client_moved_to_LTCF)
Call write_bullet_and_variable_in_case_note ("Facility name", client_moved_to_LTCF)
Call write_bullet_and_variable_in_case_note ("Facility address", LTCF_ADDR_line_01 & " " & LTCF_ADDR_line_02 & " " &  LTCF_city & " " & LTCF_state & " " & LTCF_zip_code)
IF waiver_program_change_check = 1 THEN Call write_variable_in_case_note ("* Waiver program changed from:" & waiver_program_change_from & "to" & waiver_program_change_to)
IF client_disenrolled_health_plan_check = 1 THEN Call write_variable_in_case_note ("* Client disenrolled from health plan effective" & client_disenrolled_from_healthplan)
IF new_address_check = 1 THEN Call write_variable_in_case_note ("* New Address, effective date: " & new_address_effective_date & " " & change_ADDR_line_1 & " " & change_ADDR_line_2 & " " & change_city & " " & change_state & " " & change_zip_code)
'case summary
Call write_bullet_and_variable_in_case_note ("Case actions", case_action)
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes)
If sent_5181_to_caseworker_check = 1 then Call write_variable_in_case_note("* Sent 5181 back to case manager.")
Call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure_with_error_report("Success! Please make sure your DISA and FACI panel(s) are updated if needed. Also evaluate the case for any other possible programs that can be opened, or that need to be changed or closed.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------07/21/2022
'--Tab orders reviewed & confirmed----------------------------------------------07/21/2022
'--Mandatory fields all present & Reviewed--------------------------------------07/21/2022
'--All variables in dialog match mandatory fields-------------------------------07/21/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------07/21/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------07/21/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------07/21/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/14/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------07/21/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------07/21/2022
'--PRIV Case handling reviewed -------------------------------------------------07/21/2022
'--Out-of-County handling reviewed----------------------------------------------07/21/2022----------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------07/21/2022
'--BULK - review output of statistics and run time/count (if applicable)--------07/21/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---07/21/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------07/21/2022
'--Incrementors reviewed (if necessary)-----------------------------------------07/21/2022
'--Denomination reviewed -------------------------------------------------------07/21/2022
'--Script name reviewed---------------------------------------------------------07/21/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------07/21/2022-----------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/14/2022
'--comment Code-----------------------------------------------------------------07/21/2022
'--Update Changelog for release/update------------------------------------------11/14/2022
'--Remove testing message boxes-------------------------------------------------07/21/2022
'--Remove testing code/unnecessary code-----------------------------------------07/21/2022
'--Review/update SharePoint instructions----------------------------------------11/14/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------07/21/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------07/21/2022
'--Complete misc. documentation (if applicable)---------------------------------07/21/2022
'--Update project team/issue contact (if applicable)----------------------------11/14/2022
