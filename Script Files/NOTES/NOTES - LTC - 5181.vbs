'OPTION EXPLICIT

'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - 5181.vbs"
start_time = timer

DIM beta_agency
DIM FuncLib_URL, req, fso, run_locally, default_directory

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


'Declaring variables
'DIM start_time
'DIM name_of_script
'DIM url
'DIM row
'DIM script_end_procedure
'DIM case_number_and_footer_month_dialog
'DIM case_number
'DIM footer_month
'DIM footer_year
'DIM next_month
'DIM ButtonPressed
'DIM case_note_dialog
'DIM yes_case_note_button
'DIM no_case_note_button
'DIM cancel_dialog
'DIM no_cancel_button
'DIM yes_cancel_button
'DIM MAXIS_footer_month
'DIM MAXIS_footer_year
'DIM DHS_5181_dialog_1
'DIM date_5181_editbox
'DIM date_received_editbox
'DIM lead_agency_editbox
'DIM lead_agency_assessor_editbox
'DIM casemgr_ADDR_line_01
'DIM casemgr_ADDR_line_02
'DIM casemgr_city
'DIM casemgr_state
'DIM casemgr_zip_code
'DIM phone_area_code
'DIM phone_prefix
'DIM phone_second_four
'DIM phone_extension
'DIM fax_editbox
'DIM update_SWKR_info_checkbox
'DIM update_addr_checkbox
'DIM name_of_facility_editbox
'DIM date_of_admission_editbox
'DIM facility_address_line_01
'DIM facility_address_line_02
'DIM facility_city
'DIM facility_state
'DIM facility_county_code
'DIM facility_zip_code
'DIM waiver_type_droplist
'DIM essential_community_supports_check
'DIM next_to_page_02_button
'DIM DHS_5181_dialog_2
'DIM waiver_assessment_date_editbox
'DIM needs_waiver_checkbox
'DIM estimated_effective_date_editbox
'DIM estimated_monthly_check
'DIM estimated_monthly_waiver_costs_editbox
'DIM does_not_meet_waiver_LOC_check
'DIM ongoing_waiver_case_manager_check
'DIM ongoing_waiver_case_manager_editbox
'DIM LTCF_assessment_date_editbox
'DIM meets_MALOC_check
'DIM ongoing_case_manager_check
'DIM ongoing_case_manager_editbox
'DIM ongoing_case_manager_not_available_check
'DIM does_not_meet_MALTC_LOC_check
'DIM client_applied_MA_check
'DIM client_MA_enrollee_3543_provided_check
'DIM Client_MA_enrollee_editbox
'DIM completed_3543_3531_check
'DIM completed_3543_3531_faxed_check
'DIM completed_3543_3531_faxed_editbox
'DIM please_send_3543_check
'DIM please_send_3531_check
'DIM please_send_3531_editbox
'DIM please_send_3340_check
'DIM previous_to_page_01_button
'DIM requested_1503_check
'DIM onfile_1503_check 
'DIM DHS_5181_Dialog_3
'DIM client_no_longer_meets_LOC_check
'DIM client_no_longer_meets_LOC_efffective_date_editbox
'DIM waiver_program_change_by_assessor_check
'DIM waiver_program_change_from_assessor_editbox
'DIM waiver_program_change_to_assessor_editbox
'DIM waiver_program_change_effective_date_editbox
'DIM exited_waiver_program_check
'DIM exit_waiver_end_date_editbox
'DIM client_choice_check
'DIM client_deceased_check
'DIM date_of_death_editbox
'DIM client_moved_to_LTCF_check
'DIM client_moved_to_LTCF_editbox
'DIM waiver_program_change_check
'DIM waiver_program_change_from_editbox
'DIM waiver_program_change_to_editbox
'DIM client_disenrolled_health_plan_check
'DIM client_disenrolled_from_healthplan_editbox
'DIM new_address_check
'DIM new_address_effective_date_editbox
'DIM case_action_editbox
'DIM other_notes_editbox
'DIM write_TIKL_for_worker_check
'DIM sent_5181_to_caseworker_check
'DIM worker_signature
'DIM previous_to_page_02_button
'DIM LTCF_ADDR_line_01
'DIM LTCF_ADDR_line_02
'DIM LTCF_city
'DIM LTCF_state
'DIM LTCF_county_code
'DIM LTCF_zip_code
'DIM LTCF_update_ADDR_checkbox
'DIM update_addr_new_ADDR_checkbox
'DIM change_ADDR_line_1
'DIM change_ADDR_line_2
'DIM change_city
'DIM change_state
'DIM change_county_code
'DIM change_zip_code
'DIM case_note_confirm
'DIM next_to_page_03_button
'DIM footer_month_as_date
'DIM difference_between_dates
'DIM move_on_to_case_note
'DIM from_droplist
'DIM to_droplist


'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000
 
'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
	OkButton 25, 45, 50, 15
	CancelButton 85, 45, 50, 15
EndDialog


BeginDialog DHS_5181_dialog_1, 0, 0, 361, 305, "5181 Dialog 1"
  EditBox 55, 5, 55, 15, date_5181_editbox
  EditBox 170, 5, 55, 15, date_received_editbox
  EditBox 280, 5, 70, 15, lead_agency_editbox
  EditBox 235, 30, 115, 15, lead_agency_assessor_editbox
  EditBox 65, 50, 240, 15, casemgr_ADDR_line_01
  EditBox 65, 65, 240, 15, casemgr_ADDR_line_02
  EditBox 35, 85, 80, 15, casemgr_city
  EditBox 155, 85, 40, 15, casemgr_state
  EditBox 260, 85, 45, 15, casemgr_zip_code
  EditBox 35, 105, 25, 15, phone_area_code
  EditBox 65, 105, 25, 15, phone_prefix
  EditBox 95, 105, 25, 15, phone_second_four
  EditBox 140, 105, 25, 15, phone_extension
  EditBox 190, 105, 80, 15, fax_editbox
  CheckBox 275, 105, 80, 15, "Update SWRK panel ", update_SWKR_info_checkbox
  CheckBox 60, 140, 115, 15, "Have script update ADDR panel", update_addr_checkbox
  EditBox 70, 160, 140, 15, name_of_facility_editbox
  EditBox 285, 160, 65, 15, date_of_admission_editbox
  EditBox 70, 180, 240, 15, facility_address_line_01
  EditBox 70, 195, 240, 15, facility_address_line_02
  EditBox 30, 215, 80, 15, facility_city
  EditBox 140, 215, 40, 15, facility_state
  EditBox 230, 215, 45, 15, facility_county_code
  EditBox 310, 215, 45, 15, facility_zip_code
  DropListBox 170, 250, 105, 15, "Select one..."+chr(9)+"No waiver"+chr(9)+"Alternative Care"+chr(9)+"BI diversion"+chr(9)+"BI conversion"+chr(9)+"CAC diversion"+chr(9)+"CAC conversion"+chr(9)+"CADI diversion"+chr(9)+"CADI conversion"+chr(9)+"DD diversion"+chr(9)+"DD conversion"+chr(9)+"EW diversion"+chr(9)+"EW conversion", waiver_type_droplist
  CheckBox 40, 265, 190, 10, "Essential Community Supports (DHS- 3876 is required)", essential_community_supports_check
  ButtonGroup ButtonPressed
    PushButton 250, 285, 55, 15, "Next", next_to_page_02_button
    CancelButton 310, 285, 50, 15
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


BeginDialog DHS_5181_dialog_2, 0, 0, 361, 415, "5181 Dialog 2"
  EditBox 75, 35, 45, 15, waiver_assessment_date_editbox
  EditBox 270, 50, 45, 15, estimated_effective_date_editbox
  EditBox 120, 70, 45, 15, estimated_monthly_waiver_costs_editbox
  CheckBox 5, 85, 170, 15, "Does not meet waiver services LOC requirement", does_not_meet_waiver_LOC_check
  EditBox 105, 100, 60, 15, ongoing_waiver_case_manager_editbox
  EditBox 205, 120, 45, 15, LTCF_assessment_date_editbox
  CheckBox 5, 135, 100, 10, "Meets MA-LOC requirement", meets_MALOC_check
  EditBox 125, 145, 110, 15, ongoing_case_manager_editbox
  CheckBox 5, 165, 135, 15, "Ongoing case manager not available", ongoing_case_manager_not_available_check
  CheckBox 5, 180, 115, 15, "Does not meet LOC requirement", does_not_meet_MALTC_LOC_check
  CheckBox 155, 180, 65, 10, "1503 requested?", requested_1503_check
  CheckBox 235, 180, 55, 10, "1503 on file?", onfile_1503_check
  CheckBox 5, 215, 80, 15, "Client applied for MA", client_applied_MA_check
  EditBox 240, 230, 45, 15, Client_MA_enrollee_editbox
  CheckBox 5, 245, 195, 15, "Completed DHS-3543 or DHS-3531 attached to DHS-5181", completed_3543_3531_check
  EditBox 235, 260, 45, 15, completed_3543_3531_faxed_editbox
  CheckBox 5, 275, 180, 15, "Please send DHS-3543 to client (MA enrollee)", please_send_3543_check
  EditBox 180, 290, 150, 15, please_send_3531_editbox
  CheckBox 5, 310, 205, 10, "Please send DHS-3340 to client - Asset Assessment needed", please_send_3340_check
  EditBox 235, 345, 45, 15, client_no_longer_meets_LOC_efffective_date_editbox
  DropListBox 105, 370, 60, 15, "Select one..."+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EW", from_droplist
  DropListBox 180, 370, 60, 15, "Select one..."+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EW", to_droplist
  EditBox 295, 370, 55, 15, waiver_program_change_effective_date_editbox
  ButtonGroup ButtonPressed
    PushButton 190, 400, 50, 15, "Previous", previous_to_page_01_button
    PushButton 245, 400, 50, 15, "Next", next_to_page_03_button
    CancelButton 300, 400, 50, 15
  GroupBox 0, 20, 355, 95, ""
  GroupBox 0, 115, 355, 80, ""
  GroupBox 0, 195, 355, 135, ""
  Text 0, 25, 165, 10, "**WAIVERS** Assessment date determine client:"
  Text 140, 120, 60, 15, "Assessment date:"
  Text 5, 55, 265, 10, "Needs waiver services and meets LOC. Anticipated effective date no sooner than:"
  Text 5, 335, 160, 15, "**CHANGES COMPLETED BY THE ASSESSOR**"
  Text 5, 5, 145, 10, "INITIAL REQUESTS (check all that apply):"
  Text 0, 120, 135, 15, "**LTCF** Assessment determines client: "
  Text 170, 375, 10, 10, "to:"
  Text 10, 40, 60, 10, "Assessment date:"
  Text 0, 205, 190, 10, "**MEDICAL ASSISTANCE REQUESTS/APPLICATIONS**"
  Text 245, 375, 50, 10, "Effective date:"
  GroupBox 0, 325, 355, 65, ""
  Text 10, 75, 110, 10, "Estimated monthly waiver costs:"
  Text 10, 105, 95, 10, "Ongoing case mgr assigned:"
  Text 10, 150, 110, 10, "Ongoing case manager assigned:"
  Text 10, 235, 230, 10, "Client is an MA enrollee -  If assessor provided DHS-3543, enter date:"
  Text 10, 265, 225, 10, "If completed DHS-3543 or DHS-3531 was faxed to county, enter date: "
  Text 5, 295, 170, 10, "Please send DHS-3531 to client (Not MA enrollee) at:"
  Text 5, 350, 225, 10, "Client no longer meets LOC - Effective date should be no sooner than:"
  Text 5, 375, 100, 10, "Waiver program change from:"
EndDialog


BeginDialog DHS_5181_Dialog_3, 0, 0, 361, 340, "5181 Dialog 3"
  CheckBox 5, 20, 135, 15, "Exited waiver program- Effective date: ", exited_waiver_program_check
  EditBox 140, 20, 40, 15, exit_waiver_end_date_editbox
  CheckBox 5, 35, 65, 15, "Client's choice", client_choice_check
  CheckBox 5, 50, 115, 15, "Client deceased.  Date of death:", client_deceased_check
  EditBox 135, 50, 45, 15, date_of_death_editbox
  CheckBox 5, 65, 95, 15, "Client moved to LTCF on:", client_moved_to_LTCF_check
  EditBox 135, 65, 45, 15, client_moved_to_LTCF_editbox
  CheckBox 190, 65, 115, 15, "Have script update ADDR panel", LTCF_update_ADDR_checkbox
  EditBox 65, 85, 235, 15, LTCF_ADDR_line_01
  EditBox 65, 100, 235, 15, LTCF_ADDR_line_02
  EditBox 25, 120, 55, 15, LTCF_city
  EditBox 110, 120, 45, 15, LTCF_state
  EditBox 210, 120, 45, 15, LTCF_county_code
  EditBox 295, 120, 45, 15, LTCF_zip_code
  CheckBox 5, 145, 90, 15, "Waiver program change", waiver_program_change_check
  EditBox 115, 145, 45, 15, waiver_program_change_from_editbox
  EditBox 180, 145, 45, 15, waiver_program_change_to_editbox
  CheckBox 5, 160, 175, 15, "Client disenrolled from health plan.  Effective date: ", client_disenrolled_health_plan_check
  EditBox 180, 160, 45, 15, client_disenrolled_from_healthplan_editbox
  CheckBox 5, 175, 105, 15, "New address-Effective date:", new_address_check
  EditBox 115, 175, 45, 15, new_address_effective_date_editbox
  CheckBox 190, 175, 115, 15, "Have script update ADDR panel", update_addr_new_ADDR_checkbox
  EditBox 70, 195, 235, 15, change_ADDR_line_1
  EditBox 70, 210, 235, 15, change_ADDR_line_2
  EditBox 25, 230, 60, 15, change_city
  EditBox 115, 230, 45, 15, change_state
  EditBox 215, 230, 45, 15, change_county_code
  EditBox 300, 230, 45, 15, change_zip_code
  EditBox 55, 260, 285, 15, case_action_editbox
  EditBox 55, 280, 285, 15, other_notes_editbox
  CheckBox 10, 300, 120, 10, "Inform worker of 5181 via TIKL?", write_TIKL_for_worker_check
  CheckBox 135, 300, 125, 10, "Sent 5181 back to Case Manager?", sent_5181_to_caseworker_check
  EditBox 70, 320, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 195, 320, 50, 15, "Previous", previous_to_page_02_button
    OkButton 250, 320, 50, 15
    CancelButton 305, 320, 50, 15
  Text 5, 100, 55, 15, "Address line 2:"
  Text 5, 120, 20, 15, "City:"
  Text 5, 320, 65, 15, "Worker signature:"
  Text 85, 120, 25, 15, "State:"
  Text 165, 145, 15, 15, "To: "
  Text 160, 120, 45, 15, "County code:"
  Text 260, 120, 35, 15, "Zip code:"
  Text 5, 260, 45, 15, "Case Action:"
  Text 5, 85, 60, 15, "Facility Address:"
  Text 5, 195, 60, 15, "Address line 1:"
  Text 5, 210, 55, 15, "Address line 2:"
  Text 5, 230, 20, 15, "City:"
  Text 90, 230, 25, 15, "State:"
  Text 5, 5, 125, 15, "**CHANGES** (check all that apply):"
  Text 165, 230, 45, 15, "County code:"
  Text 95, 145, 20, 15, "From:"
  Text 265, 230, 35, 15, "Zip code:"
  Text 5, 280, 45, 15, "Other notes:"
  GroupBox 0, 0, 355, 250, ""
EndDialog


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

'Grabbing the case number
call MAXIS_case_number_finder(case_number)

'Grabbing the footer month/year
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
	footer_month = MAXIS_footer_month
	call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
	If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number
Do
	Dialog case_number_and_footer_month_dialog
	If ButtonPressed = 0 then stopscript
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
transmit

'Dialog completed by worker. Each dialog follows this process:
'  1. Show the dialog and validate that next/OK or prev is pressed
'  2. Do the validation on that page, but contain a "if ButtonPressed = prev then exit do" to skip the validation if previous is pressed
'  3. Validate again that next/OK or prev is pressed
'  4. Loop until next is pressed, which will loop back to the previous dialog.
Do
	Do
		Do
			Do
				Dialog DHS_5181_dialog_1			'Displays the first dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.	
				'MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
			Loop until ButtonPressed = next_to_page_02_button
			IF waiver_type_droplist = "Select one..." THEN MsgBox "Choose waiver type (or select 'no waiver')."		'Requires the user to select a waiver
		Loop until waiver_type_droplist <> "Select one..."
	Loop until ButtonPressed = next_to_page_02_button
	Do
		Do
			Do
				Do
					Dialog DHS_5181_dialog_2			'Displays the second dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					'MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
				If ButtonPressed = previous_to_page_01_button THEN exit do
				If (from_droplist = "Select one..." AND to_droplist <> "Select one...") OR (from_droplist <> "Select one..." AND to_droplist = "Select one...") THEN Msgbox	"You must enter valid selections for the waiver program change 'to' and 'from'." 'Requires the user to enter a droplist item
			Loop until (from_droplist = "Select one..." AND to_droplist = "Select one...") OR (from_droplist <> "Select one..." AND to_droplist <> "Select one...") 'Loops until both from and to are filled out, or neither.
		Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
		If ButtonPressed = previous_to_page_01_button then exit do
		Do
			Do				
				Do
					Dialog DHS_5181_Dialog_3			'Displays the third dialog
					cancel_confirmation					'Asks if you're sure you want to cancel, and cancels if you select that.
					'MAXIS_dialog_navigation				'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button
				If ButtonPressed = previous_to_page_02_button THEN exit do
				If case_action_editbox = "" or worker_signature = "" OR (exited_waiver_program_check = checked AND exit_waiver_end_date_editbox = "") OR _
					(client_deceased_check =  checked AND date_of_death_editbox = "") OR (client_moved_to_LTCF_check = checked AND client_moved_to_LTCF_editbox = "") OR _
					(waiver_program_change_check = checked AND waiver_program_change_from_editbox = "" AND waiver_program_change_to_editbox = "") OR _
					(client_disenrolled_health_plan_check = checked AND client_disenrolled_from_healthplan_editbox = "") OR (new_address_check = checked AND new_address_effective_date_editbox =  "") THEN
					MsgBox "You need to:" & vbNewLine & vbNewLine & _
					"-Complete a field next to an option that was checked, and/or" & vbNewLine & _	
					"-Case note the 'case actions' section, and/or" & vbNewLine & _
					"-Sign your case note." & vbNewLine & vbNewLine & _
					"Check these items after pressing ''OK''."	
				ELSE
					move_on_to_case_note = True			'If all of that stuff is fine, this will set to true and the loop can end.
				End if
			Loop until move_on_to_case_note = True
		Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button
	Loop until ButtonPressed = -1
	CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
Loop until case_note_confirm = TRUE		




'Dollar bill symbol will be added to numeric variables 
IF estimated_monthly_waiver_costs_editbox <> "" THEN estimated_monthly_waiver_costs_editbox = "$" & estimated_monthly_waiver_costs_editbox

'Checking to see that we're in MAXIS
call check_for_MAXIS(False)

'ACTIONS----------------------------------------------------------------------------------------------------

'Inform worker of 5181 via TIKL (check box selected)
IF write_TIKL_for_worker_check = 1 THEN 
	'Go to DAIL/WRIT
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	
	'Writes TIKL to worker
	call write_variable_in_TIKL("A DHS 5181 has been received for this case.  Please review the case and case notes.")
	transmit
	PF3
END If

'Updates STAT MEMB with client's date of death (client_deceased_check)
IF client_deceased_check = 1 THEN 
	'Go to STAT MEMB
	
	'Creates a new variable with footer_month and footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = footer_month & "/01/" & footer_year
	
	'Calculates the difference between the two dates (date of death and footer month)
	difference_between_dates = DateDiff("m", date_of_death_editbox, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		Call convert_date_into_MAXIS_footer_month(date_of_death_editbox, footer_month, footer_year)
		EMWriteScreen footer_month, 20, 43
		EMWriteScreen footer_year, 20, 46
		Transmit
	END IF
	
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	PF9
		
	'Writes in DOD from the date_of_death_editbox
	Call create_MAXIS_friendly_date_with_YYYY(date_of_death_editbox, 0, 19, 42)	
	transmit
	PF3
	transmit
END IF

'------ADDRESS UPDATES----------------------------------------------------------------------------------------------------
'Updates ADDR if selected on DIALOG 1 "have script update ADDR panel"
IF update_addr_checkbox = 1 THEN 
	'Creates a new variable with footer_month and footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = footer_month & "/01/" & footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", date_of_admission_editbox, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(date_of_admission_editbox, footer_month, footer_year)
		EMWriteScreen footer_month, 20, 43
		EMWriteScreen footer_year, 20, 46
		Transmit
	END IF
	
	'Go to STAT/ADDR
	Call navigate_to_MAXIS_screen("STAT", "ADDR")

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "______", 4, 43 
	EMWriteScreen "______________________", 6, 43
	EMWriteScreen "______________________", 7, 43
	EMWriteScreen "_______________", 8, 43
	EMWriteScreen "__", 8, 66
	EMWriteScreen "__", 9, 66
	EMWriteScreen "_____", 9, 43
	
	'Writes in the new info
	Call Create_MAXIS_friendly_date(date_of_admission_editbox, 0, 4, 43)
	EMWriteScreen facility_address_line_01, 6, 43
	EMWriteScreen facility_address_line_02, 7, 43
	EMWriteScreen facility_city, 8, 43
	EMWriteScreen facility_state, 8, 66
	EMWriteScreen facility_county_code, 9, 66
	EMWriteScreen facility_zip_code, 9, 43
	
	transmit
	transmit
	transmit
END If


'Updates ADDR if selected on DIALOG 3 "have script update ADDR panel" for move to LTCF
IF LTCF_update_ADDR_checkbox = 1 THEN 
		'Creates a new variable with footer_month and footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = footer_month & "/01/" & footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", client_moved_to_LTCF_editbox, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(client_moved_to_LTCF_editbox, footer_month, footer_year)
		EMWriteScreen footer_month, 20, 43
		EMWriteScreen footer_year, 20, 46
		Transmit
	END IF
	
	'Go to STAT/ADDR
	Call navigate_to_MAXIS_screen("STAT", "ADDR")

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "______", 4, 43 
	EMWriteScreen "______________________", 6, 43
	EMWriteScreen "______________________", 7, 43
	EMWriteScreen "_______________", 8, 43
	EMWriteScreen "__", 8, 66
	EMWriteScreen "__", 9, 66
	EMWriteScreen "_____", 9, 43
	
	'Writes in the new info
	Call Create_MAXIS_friendly_date(client_moved_to_LTCF_editbox, 0, 4, 43)
	EMWriteScreen LTCF_ADDR_line_01, 6, 43
	EMWriteScreen LTCF_ADDR_line_02, 7, 43
	EMWriteScreen LTCF_city, 8, 43
	EMWriteScreen LTCF_state, 8, 66
	EMWriteScreen LTCF_county_code, 9, 66
	EMWriteScreen LTCF_zip_code, 9, 43
	
	transmit
	transmit
	transmit
END If


'Updates ADDR if selected on DIALOG 3 "have script update ADDR panel" for new address
IF update_addr_new_ADDR_checkbox = 1 THEN 
	'Creates a new variable with footer_month and footer_year concatenated into a single date starting on the 1st of the month.
	footer_month_as_date = footer_month & "/01/" & footer_year

	'Calculates the difference between the two dates (date of admission and footer month)
	difference_between_dates = DateDiff("m", new_address_effective_date_editbox, footer_month_as_date)

	'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
	If difference_between_dates <> 0 THEN
		back_to_SELF
		CALL convert_date_into_MAXIS_footer_month(new_address_effective_date_editbox, footer_month, footer_year)
		EMWriteScreen footer_month, 20, 43
		EMWriteScreen footer_year, 20, 46
		Transmit
	END IF
	
	'Go to STAT/ADDR
	Call navigate_to_MAXIS_screen("STAT", "ADDR")

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "______", 4, 43 
	EMWriteScreen "______________________", 6, 43
	EMWriteScreen "______________________", 7, 43
	EMWriteScreen "_______________", 8, 43
	EMWriteScreen "__", 8, 66
	EMWriteScreen "__", 9, 66
	EMWriteScreen "_____", 9, 43
	
	'Writes in the new info
	Call Create_MAXIS_friendly_date(new_address_effective_date_editbox, 0, 4, 43)
	EMWriteScreen change_ADDR_line_1, 6, 43
	EMWriteScreen change_ADDR_line_2, 7, 43
	EMWriteScreen change_city, 8, 43
	EMWriteScreen change_state, 8, 66
	EMWriteScreen change_county_code, 9, 66
	EMWriteScreen change_zip_code, 9, 43
	
	transmit
	transmit
	transmit
END If


'Updates SWKR panel with Name, address and phone number if checked on DIALOG 1
If update_SWKR_info_checkbox = 1 THEN
	'Go to STAT/SWKR
	Call navigate_to_MAXIS_screen("STAT", "SWKR")

	'creates a new panel if one doesn't exist, and will needs new if there is not one
	Call Create_panel_if_nonexistent	

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "___________________________________", 6, 32
	EMWriteScreen "______________________", 8, 32
	EMWriteScreen "______________________", 9, 32
	EMWriteScreen "_______________", 10, 32
	EMWriteScreen "__", 10, 54
	EMWriteScreen "_____", 10,63
	EMWriteScreen "___", 12, 34
	EMWriteScreen "___", 12, 40
	EMWriteScreen "____", 12, 44
	EMWriteScreen "____", 12, 54
	
	'Writes in the new info into the SWKR panel
	EMWriteScreen lead_agency_assessor_editbox, 6, 32
	EMWriteScreen casemgr_ADDR_line_01, 8, 32
	EMWriteScreen casemgr_ADDR_line_02, 9, 32
	EMWriteScreen casemgr_city, 10, 32
	EMWriteScreen casemgr_state, 10, 54
	EMWriteScreen casemgr_zip_code, 10, 63
	EMWriteScreen phone_area_code, 12, 34
	EMWriteScreen phone_prefix, 12, 40
	EMWriteScreen phone_second_four, 12, 44
	EMWriteScreen phone_extension, 12, 54
	EMWriteScreen "Y", 15, 63
	
	transmit
	transmit
	PF3
END IF	

'Updates SWKR panel with ongoing case manager assigned
If ongoing_waiver_case_manager_check = 1 THEN
	'Go to STAT/SWKR
	Call navigate_to_MAXIS_screen("STAT", "SWKR")

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "___________________________________", 6, 32
	
	'Writes in new case manager name
	EMWriteScreen ongoing_waiver_case_manager_editbox, 6, 32
	
	transmit
	transmit
	PF3
END IF
	
'Updates SWKR panel with ongoing case manager assigned
If ongoing_case_manager_check = 1 THEN
	'Go to STAT/SWKR
	Call navigate_to_MAXIS_screen("STAT", "SWKR")

	'Go into edit mode
	PF9

	'Blanks out the old info
	EMWriteScreen "___________________________________", 6, 32
	
	'Writes in new case manager name
	EMWriteScreen ongoing_case_manager_editbox, 6, 32
	
	transmit
	transmit
	PF3
END IF
	
'Checking to see that we're in MAXIS
call check_for_MAXIS(True)

'function to navigate user to case note and make a new one
Call start_a_blank_CASE_NOTE

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Information from DHS 5181 Dialog 1
'Contact information
Call write_variable_in_case_note ("~~~DHS 5181 rec'd~~~")											
Call write_bullet_and_variable_in_case_note ("Date of 5181", date_5181_editbox )		
Call write_bullet_and_variable_in_case_note ("Date received", date_received_editbox)			
Call write_bullet_and_variable_in_case_note ("Lead Agency", lead_agency_editbox)
Call write_bullet_and_variable_in_case_note ("Lead Agency Assessor/Case Manager",lead_agency_assessor_editbox)		  				 
Call write_bullet_and_variable_in_case_note ("Address", casemgr_ADDR_line_01 & " " & casemgr_ADDR_line_02 & " " & casemgr_city & " " & casemgr_state & " " & casemgr_zip_code)					 
Call write_bullet_and_variable_in_case_note ("Phone", phone_area_code & " " & phone_prefix & " " & phone_second_four & " " & phone_extension)
Call write_bullet_and_variable_in_case_note ("Fax", fax_editbox)
'STATUS
Call write_bullet_and_variable_in_case_note ("Name of Facility", name_of_facility_editbox)											
Call write_bullet_and_variable_in_case_note ("Date of admission", date_of_admission_editbox)		
Call write_bullet_and_variable_in_case_note ("Facility address", facility_address_line_01 & " " & facility_address_line_02 & " " & facility_city & " " & facility_state & " " & facility_zip_code)	
IF waiver_type_droplist <> "No waiver" then call write_bullet_and_variable_in_case_note("Client is requesting services/enrolled in waiver type", waiver_type_droplist)	
IF essential_community_supports_check = 1 THEN Call write_variable_in_case_note ("* Essential Community supports.  Client does not meet LOC requirements.")	

'Information from DHS 5181 Dialog 2
'Waivers
Call write_bullet_and_variable_in_case_note ("Waiver Assessment Date", waiver_assessment_date_editbox)	
Call write_bullet_and_variable_in_case_note ("Assessment determines that client needs waiver services and meets LOC requirements.  Anticipated effective date no sooner than", estimated_effective_date_editbox)
Call write_bullet_and_variable_in_case_note ("Estimated monthly waiver costs", estimated_monthly_waiver_costs_editbox)
IF does_not_meet_waiver_LOC_check = 1 THEN Call write_variable_in_case_note ("* Client does not meet LOC requirements for waivered services.")
Call write_bullet_and_variable_in_case_note ("Ongoing case manager is", ongoing_waiver_case_manager_editbox)
'LTCF
Call write_bullet_and_variable_in_case_note ("LTCF Assessment Date", LTCF_assessment_date_editbox)	
IF meets_MALOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client meets the LOC requirement")
Call write_bullet_and_variable_in_case_note("Ongoing case manager is", ongoing_case_manager_editbox)
IF ongoing_case_manager_not_available_check = 1 THEN Call write_variable_in_case_note ("* Ongoing Case Manager not available")
IF does_not_meet_MALTC_LOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client does not meet LOC requirements for LTCF's.")
IF requested_1503_check = 1 THEN Call write_variable_in_case_note ("* A DHS-1503 has been requested from the LTCF.")
IF onfile_1503_check = 1 THEN Call write_variable_in_case_note ("A DHS-1503 has been provided.")
'MA requests/applications
IF client_applied_MA_check = 1 THEN Call write_variable_in_case_note ("* Client has applied for MA")
Call write_bullet_and_variable_in_case_note ("Client is an MA enrollee. Assessor provided a DHS-3543 on", Client_MA_enrollee_editbox)
IF completed_3543_3531_check = 1 THEN Call write_variable_in_case_note ("* Completed DHS-3543 or DHS-3531 attached to DHS 5181")
Call write_bullet_and_variable_in_case_note ("Completed DHS-3543 or DHS-3531 faxed to county on", completed_3543_3531_faxed_editbox)
IF please_send_3543_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested that a DHS-3543 be sent to the MA enrollee or AREP.")
Call write_bullet_and_variable_in_case_note ("* Case manager has requested that a DHS-3531 be sent to a non-MA enrollee at", please_send_3531_editbox)
IF please_send_3340_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested an Asset Assessment, DHS 3340, be send to the client or AREP")
'changes completed by the assessor
Call write_bullet_and_variable_in_case_note ("Client no longer meets LOC - Effective date should be no sooner than", client_no_longer_meets_LOC_efffective_date_editbox)					 
IF from_droplist <> "Select one..." AND to_droplist <> "Select one.." THEN Call write_bullet_and_variable_in_case_note ("Waiver program changed from", from_droplist & " to: " & to_droplist & ". Effective date: " & waiver_program_change_effective_date_editbox)

'Information from DHS 5181 Dialog 3
'changes
IF exited_waiver_program_check = 1 THEN Call write_variable_in_case_note("* Exited waiver program.  Effective date: " & exit_waiver_end_date_editbox)
IF client_choice_check = 1 THEN Call write_variable_in_case_note ("* Client has chosen to exit the waiver program")
IF client_deceased_check = 1 THEN Call write_variable_in_case_note ("* Client is deceased.  Date of death: " & date_of_death_editbox)
IF client_moved_to_LTCF_check = 1 THEN Call write_variable_in_case_note ("* Client moved to LTCF on" & client_moved_to_LTCF_editbox)
Call write_bullet_and_variable_in_case_note ("Facility name", client_moved_to_LTCF_editbox)
Call write_bullet_and_variable_in_case_note ("Facility address", LTCF_ADDR_line_01 & " " & LTCF_ADDR_line_02 & " " &  LTCF_city & " " & LTCF_state & " " & LTCF_zip_code)
IF waiver_program_change_check = 1 THEN Call write_variable_in_case_note ("* Waiver program changed from:" & waiver_program_change_from_editbox & "to" & waiver_program_change_to_editbox)
IF client_disenrolled_health_plan_check = 1 THEN Call write_variable_in_case_note ("* Client disenrolled from health plan effective" & client_disenrolled_from_healthplan_editbox)
IF new_address_check = 1 THEN Call write_variable_in_case_note ("* New Address, effective date: " & new_address_effective_date_editbox & " " & change_ADDR_line_1 & " " & change_ADDR_line_2 & " " & change_city & " " & change_state & " " & change_zip_code)
'case summary
Call write_bullet_and_variable_in_case_note ("Case actions", case_action_editbox)
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes_editbox)
Call write_variable_in_case_note ("---")						 
call write_variable_in_case_note (worker_signature)
MsgBox "Make sure your DISA and FACI panel(s) are updated if needed. Please also evaluate the case for any other possible programs that can be opened, or that need to be changed or closed."

script_end_procedure("")
