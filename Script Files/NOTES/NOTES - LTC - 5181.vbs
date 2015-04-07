OPTION EXPLICIT

'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - 5181.vbs"
start_time = timer

'FUNCTIONS LIBRARY
'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER FUNCTIONS LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'Declaring variables
DIM start_time
DIM name_of_script
DIM url
DIM req
DIM fso
DIM row
DIM script_end_procedure
DIM case_number_and_footer_month_dialog
DIM case_number
DIM footer_month
DIM footer_year
DIM next_month
DIM ButtonPressed
DIM case_note_dialog
DIM yes_case_note_button
DIM no_case_note_button
DIM cancel_dialog
DIM no_cancel_button
DIM yes_cancel_button
DIM MAXIS_footer_month
DIM MAXIS_footer_year
DIM DHS_5181_dialog_1
DIM date_5181_editbox
DIM date_received_editbox
DIM lead_agency_editbox
DIM lead_agency_assessor_editbox
DIM casemgr_ADDR_line_01
DIM casemgr_ADDR_line_02
DIM casemgr_city
DIM casemgr_state
DIM casemgr_zip_code
DIM phone_area_code
DIM phone_prefix
DIM phone_second_four
DIM phone_extension
DIM fax_editbox
DIM update_SWKR_info_checkbox
DIM update_addr_checkbox
DIM name_of_facility_editbox
DIM date_of_admission_editbox
DIM facility_address_line_01
DIM facility_address_line_02
DIM facility_city
DIM facility_state
DIM facility_county_code
DIM facility_zip_code
DIM AC_check
DIM BI_check
DIM CAC_check
DIM CADI_check
DIM DD_check
DIM EW_check
DIM diversion_check
DIM conversion_check
DIM essential_community_supports_check
DIM next_to_page_02_button
DIM DHS_5181_dialog_2
DIM waiver_assessment_date_editbox
DIM needs_waiver_check
DIM estimated_effective_date_editbox
DIM estimated_monthly_check
DIM estimated_monthly_waiver_costs_editbox
DIM does_not_meet_waiver_LOC_check
DIM ongoing_waiver_case_manager_check
DIM ongoing_waiver_case_manager_editbox
DIM LTCF_assessment_date_editbox
DIM meets_MALOC_check
DIM ongoing_case_manager_check
DIM ongoing_case_manager_editbox
DIM ongoing_case_manager_not_available_check
DIM does_not_meet_MALTC_LOC_check
DIM client_applied_MA_check
DIM client_MA_enrollee_3543_provided_check
DIM Client_MA_enrollee_editbox
DIM completed_3543_3531_check
DIM completed_3543_3531_faxed_check
DIM completed_3543_3531_faxed_editbox
DIM please_send_3543_check
DIM please_send_3531_check
DIM please_send_3531_editbox
DIM please_send_3340_check
DIM previous_to_page_01_button
DIM requested_1503_check
DIM onfile_1503_check 
DIM next_to_page_03_button
DIM DHS_5181_Dialog_3
DIM client_no_longer_meets_LOC_check
DIM client_no_longer_meets_LOC_efffective_date_editbox
DIM waiver_program_change_by_assessor_check
DIM waiver_program_change_from_assessor_editbox
DIM waiver_program_change_to_assessor_editbox
DIM waiver_program_change_effective_date_editbox
DIM exited_waiver_program_check
DIM exit_waiver_end_date_editbox
DIM client_choice_check
DIM client_deceased_check
DIM date_of_death_editbox
DIM client_moved_to_LTCF_check
DIM client_moved_to_LTCF_editbox
DIM waiver_program_change_check
DIM waiver_program_change_from_editbox
DIM waiver_program_change_to_editbox
DIM client_disenrolled_health_plan_check
DIM client_disenrolled_from_healthplan_editbox
DIM new_address_check
DIM new_address_effective_date_editbox
DIM case_action_editbox
DIM other_notes_editbox
DIM write_TIKL_for_worker_check
DIM sent_5181_to_caseworker_check
DIM worker_signature
DIM previous_to_page_02_button
DIM LTCF_ADDR_line_01
DIM LTCF_ADDR_line_02
DIM LTCF_city
DIM LTCF_state
DIM LTCF_county_code
DIM LTCF_zip_code
DIM LTCF_update_ADDR_checkbox
DIM update_addr_new_ADDR_checkbox
DIM change_ADDR_line_1
DIM change_ADDR_line_2
DIM change_city
DIM change_state
DIM change_county_code
DIM change_zip_code
DIM case_note_confirm

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


BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog


BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
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
  CheckBox 10, 250, 20, 10, "AC", AC_check
  CheckBox 35, 250, 20, 10, "BI", BI_check
  CheckBox 60, 250, 30, 10, "CAC", CAC_check
  CheckBox 90, 250, 30, 10, "CADI", CADI_check
  CheckBox 125, 250, 25, 10, "DD", DD_check
  CheckBox 155, 250, 25, 10, "EW", EW_check
  CheckBox 250, 250, 40, 10, "Diversion", diversion_check
  CheckBox 300, 250, 50, 10, "Conversion", conversion_check
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
  Text 195, 250, 45, 10, "Choose one:"
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


BeginDialog DHS_5181_dialog_2, 0, 0, 361, 340, "5181 Dialog 2"
  EditBox 250, 20, 45, 15, waiver_assessment_date_editbox
  CheckBox 5, 35, 315, 10, "Needs waiver services and meets LOC requirement", needs_waiver_check
  EditBox 160, 50, 45, 15, estimated_effective_date_editbox
  CheckBox 5, 65, 115, 15, "Estimated monthly waiver costs:", estimated_monthly_check
  EditBox 125, 65, 45, 15, estimated_monthly_waiver_costs_editbox
  CheckBox 5, 80, 170, 15, "Does not meet waiver services LOC requirement", does_not_meet_waiver_LOC_check
  CheckBox 5, 95, 105, 15, "Ongoing case mgr assigned:", ongoing_waiver_case_manager_check
  EditBox 110, 95, 60, 15, ongoing_waiver_case_manager_editbox
  EditBox 210, 115, 45, 15, LTCF_assessment_date_editbox
  CheckBox 5, 125, 100, 15, "Meets MA-LOC requirement", meets_MALOC_check
  CheckBox 5, 145, 120, 15, "Ongoing case manager assigned:", ongoing_case_manager_check
  EditBox 130, 145, 110, 15, ongoing_case_manager_editbox
  CheckBox 5, 160, 135, 15, "Ongoing case manager not available", ongoing_case_manager_not_available_check
  CheckBox 5, 175, 115, 15, "Does not meet LOC requirement", does_not_meet_MALTC_LOC_check
  CheckBox 0, 210, 80, 15, "Client applied for MA", client_applied_MA_check
  CheckBox 0, 225, 205, 15, "Client is an MA enrollee - Assessor provided DHS-3543 on:", client_MA_enrollee_3543_provided_check
  EditBox 205, 230, 45, 15, Client_MA_enrollee_editbox
  CheckBox 0, 240, 155, 15, "Completed DHS-3543 or DHS-3531 attached", completed_3543_3531_check
  CheckBox 0, 255, 190, 15, "Completed DHS-3543 or DHS-3531 faxed to county on: ", completed_3543_3531_faxed_check
  EditBox 190, 255, 45, 15, completed_3543_3531_faxed_editbox
  CheckBox 0, 270, 180, 15, "Please send DHS-3543 to client (MA enrollee)", please_send_3543_check
  CheckBox 0, 285, 185, 15, "Please send DHS-3531 to client (Not MA enrollee) at:", please_send_3531_check
  EditBox 190, 285, 150, 15, please_send_3531_editbox
  CheckBox 0, 300, 205, 15, "Please send DHS-3340 to client - Asset Assessment needed", please_send_3340_check
  CheckBox 30, 325, 65, 10, "1503 requested?", requested_1503_check
  CheckBox 105, 325, 55, 10, "1503 on file?", onfile_1503_check
  ButtonGroup ButtonPressed
    PushButton 195, 320, 50, 15, "Previous", previous_to_page_01_button
    PushButton 250, 320, 50, 15, "Next", next_to_page_03_button
    CancelButton 305, 320, 50, 15
  Text 5, 195, 190, 15, "**MEDICAL ASSISTANCE REQUESTS/APPLICATIONS**"
  Text 5, 0, 145, 15, "INITIAL REQUESTS (check all that apply):"
  Text 5, 115, 135, 15, "**LTCF** Assessment determines client: "
  Text 185, 20, 60, 15, "Assessment date:"
  GroupBox 0, 15, 355, 95, ""
  GroupBox 0, 110, 355, 80, ""
  GroupBox 0, 190, 355, 125, ""
  Text 5, 20, 165, 15, "**WAIVERS** Assessment date determine client:"
  Text 145, 115, 60, 15, "Assessment date:"
  Text 20, 50, 135, 15, "Anticipated effective date no sooner than:"
EndDialog


BeginDialog DHS_5181_Dialog_3, 0, 0, 361, 415, "5181 Dialog 3"
  CheckBox 0, 25, 235, 15, "Client no longer meets LOC - Effective date should be no sooner than: ", client_no_longer_meets_LOC_check
  EditBox 240, 25, 45, 15, client_no_longer_meets_LOC_efffective_date_editbox
  CheckBox 0, 45, 90, 10, "Waiver program change", waiver_program_change_by_assessor_check
  EditBox 115, 45, 45, 15, waiver_program_change_from_assessor_editbox
  EditBox 180, 45, 45, 15, waiver_program_change_to_assessor_editbox
  EditBox 285, 45, 45, 15, waiver_program_change_effective_date_editbox
  CheckBox 5, 90, 135, 15, "Exited waiver program- Effective date: ", exited_waiver_program_check
  EditBox 140, 90, 40, 15, exit_waiver_end_date_editbox
  CheckBox 5, 105, 65, 15, "Client's choice", client_choice_check
  CheckBox 5, 120, 115, 15, "Client deceased.  Date of death:", client_deceased_check
  EditBox 125, 120, 45, 15, date_of_death_editbox
  CheckBox 5, 135, 95, 15, "Client moved to LTCF on:", client_moved_to_LTCF_check
  EditBox 125, 135, 45, 15, client_moved_to_LTCF_editbox
  CheckBox 185, 135, 120, 15, "Have script update ADDR panel", LTCF_update_ADDR_checkbox
  EditBox 65, 155, 235, 15, LTCF_ADDR_line_01
  EditBox 65, 170, 235, 15, LTCF_ADDR_line_02
  EditBox 25, 190, 60, 15, LTCF_city
  EditBox 115, 190, 45, 15, LTCF_state
  EditBox 220, 190, 45, 15, LTCF_county_code
  EditBox 305, 190, 45, 15, LTCF_zip_code
  CheckBox 0, 215, 90, 15, "Waiver program change", waiver_program_change_check
  EditBox 115, 215, 45, 15, waiver_program_change_from_editbox
  EditBox 180, 215, 45, 15, waiver_program_change_to_editbox
  CheckBox 0, 230, 175, 15, "Client disenrolled from health plan.  Effective date: ", client_disenrolled_health_plan_check
  EditBox 180, 230, 45, 15, client_disenrolled_from_healthplan_editbox
  CheckBox 0, 245, 105, 15, "New address-Effective date:", new_address_check
  EditBox 110, 245, 45, 15, new_address_effective_date_editbox
  EditBox 70, 265, 235, 15, change_ADDR_line_1
  EditBox 70, 280, 235, 15, change_ADDR_line_2
  EditBox 30, 300, 60, 15, change_city
  EditBox 120, 300, 45, 15, change_state
  EditBox 225, 300, 45, 15, change_county_code
  EditBox 310, 300, 45, 15, change_zip_code
  EditBox 55, 330, 285, 15, case_action_editbox
  EditBox 55, 350, 285, 15, other_notes_editbox
  CheckBox 65, 370, 120, 10, "Inform worker of 5181 via TIKL?", write_TIKL_for_worker_check
  CheckBox 195, 370, 125, 10, "Sent 5181 back to Case Manager?", sent_5181_to_caseworker_check
  EditBox 70, 390, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 195, 390, 50, 15, "Previous", previous_to_page_02_button
    OkButton 250, 390, 50, 15
    CancelButton 305, 390, 50, 15
  Text 5, 350, 45, 15, "Other notes:"
  Text 0, 170, 55, 15, "Address line 2:"
  Text 5, 10, 165, 15, "**CHANGES COMPLETED BY THE ASSESSOR**"
  Text 5, 190, 20, 15, "City:"
  Text 5, 390, 65, 15, "Worker signature:"
  Text 90, 190, 25, 15, "State:"
  Text 165, 215, 15, 15, "To: "
  Text 170, 190, 45, 15, "County code:"
  GroupBox 0, 0, 350, 70, ""
  Text 270, 190, 35, 15, "Zip code:"
  Text 5, 330, 45, 15, "Case Action:"
  Text 0, 155, 60, 15, "Facility Address:"
  Text 5, 265, 60, 15, "Address line 1:"
  Text 165, 45, 15, 15, "To: "
  Text 5, 280, 55, 15, "Address line 2:"
  Text 235, 45, 50, 15, "Effective Date:"
  Text 10, 300, 20, 15, "City:"
  Text 95, 45, 20, 15, "From:"
  Text 95, 300, 25, 15, "State:"
  Text 5, 75, 125, 15, "**CHANGES** (check all that apply):"
  Text 175, 300, 45, 15, "County code:"
  Text 95, 215, 20, 15, "From:"
  Text 275, 300, 35, 15, "Zip code:"
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

'Dialog completed by worker
Do
	Do
		Do
			Dialog DHS_5181_dialog_1			'Displays the first dialog
			cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.	
			MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
		Loop until ButtonPressed = next_to_page_02_button
		Do
			Do
				Dialog DHS_5181_dialog_2			'Displays the second dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
			Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button		'If you press either the next or previous button, this loop ends
			If ButtonPressed = previous_to_page_01_button then exit do		'If the button was previous, it exits this do loop and is caught in the next one, which sends you back to Dialog 1 because of the "If ButtonPressed = previous_to_page_01_button then exit do" later on
			Do
				Dialog DHS_5181_Dialog_3			'Displays the third dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				If ButtonPressed = previous_to_page_02_button then exit do		'Exits this do...loop here if you press previous. The second ""loop until ButtonPressed = -1" gets caught, and it loops back to the "Do" after "Loop until ButtonPressed = next_to_page_02_button"
			Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button		'If OK or PREV, it exits the loop here, which is weird because the above also causes it to exit
		Loop until ButtonPressed = -1	'Because this is in here a second time, it triggers a return to the "Dialog CAF_dialog_02" line, where all those "DOs" start again!!!!!
		If ButtonPressed = previous_to_page_01_button then exit do 	'This exits this particular loop again for prev button on page 2, which sends you back to page 1!!
		If case_action_editbox = "" or worker_signature = "" THEN 'Tells the worker what's required in a MsgBox.
			MsgBox "You need to:" & chr(13) & chr(13) & _
			  "-Case actions section, and/or" & chr(13) & _
			  "-Sign your case note." & chr(13) & chr(13) & _
			  "Check these items after pressing ''OK''."	
		End if
	Loop until case_action_editbox <> ""  and worker_signature <> ""	'Loops all of that until those four sections are finished. Let's move that over to those particular pages. Folks would be less angry that way I bet.
	
	CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
									  

'Dollar bill symbol will be added to numeric variables 
IF estimated_monthly_waiver_costs_editbox <> "" THEN estimated_monthly_waiver_costs_editbox = "$" & estimated_monthly_waiver_costs_editbox

'Checking to see that we're in MAXIS
call check_for_MAXIS(True)

'ACTIONS----------------------------------------------------------------------------------------------------

'Inform worker of 5181 via TIKL (check box selected)
IF write_TIKL_for_worker_check = 1 THEN 
	'Go to DAIL/WRIT
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	
	'Writes TIKL to worker
	EMWriteScreen "A DHS 5181 has been received for this case.  Please review case and case note", 9, 3
	
	transmit
	PF3
END If

'Updates STAT MEMB with client's date of death (client_deceased_check)
IF client_deceased_check = 1 THEN 
	'Go to STAT MEMB
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	PF9
	
	'Writes in DOD from the date_of_death_editbox
	EMWriteScreen date_of_death_editbox, 19, 42	
	
	transmit
	PF3
	tranmit
END If


'Updates ADDR if selected on DIALOG 1 "have script update ADDR panel"
IF update_addr_checkbox = 1 THEN 
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
	EMWriteScreen date_of_admission_editbox, 4, 43
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
	EMWriteScreen client_moved_to_LTCF_editbox, 4, 43
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
	EMWriteScreen new_address_effective_date_editbox, 4, 43
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


'creates a new panel if one doesn't exist, and will needs new if there is not one
call Create_panel_if_nonexistant	

'Updates SWKR panel with Name, address and phone number if checked
If update_SWKR_info_checkbox = 1 THEN
	'Go to STAT/SWKR
	Call navigate_to_MAXIS_screen("STAT", "SWKR")

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

'Checking to see that we're in MAXIS
call check_for_MAXIS(True)

'function to navigate user to case note
Call navigate_to_screen ("case", "note")						
PF9	

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Information from DHS 5181 Dialog 1
Call write_variable_in_case_note ("~~~DHS 5181 rec'd~~~")											
Call write_bullet_and_variable_in_case_note ("Date of 5181", date_5181_editbox )		
Call write_bullet_and_variable_in_case_note ("Date received", date_received_editbox)			
Call write_bullet_and_variable_in_case_note ("Lead Agency", lead_agency_editbox)
Call write_bullet_and_variable_in_case_note ("Lead Agency Assessor/Case Manager",lead_agency_assessor_editbox)		  				 
Call write_bullet_and_variable_in_case_note ("Address", casemgr_ADDR_line_01 & casemgr_ADDR_line_02 & casemgr_city & casemgr_state & casemgr_zip_code)					 
Call write_bullet_and_variable_in_case_note ("Phone", phone_area_code & phone_prefix & phone_second_four & phone_extension)
Call write_bullet_and_variable_in_case_note ("Fax", fax_editbox)
Call write_bullet_and_variable_in_case_note ("Name of Facility", name_of_facility_editbox)											
Call write_bullet_and_variable_in_case_note ("Date of admission", date_of_admission_editbox)		
Call write_bullet_and_variable_in_case_note ("Facility address", facility_address_line_01 & facility_address_line_02 & facility_city & facility_state & facility_zip_code)	
'OR
IF AC_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: AC")		
IF BI_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: BI")	
IF CAC_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: CAC")	
IF CADI_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: CADI")	
IF DD_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: DD")	
IF EW_check = 1 THEN Call write_variable_in_case_note ("* The client is currently requesting services/enrolled in the following waiver program: EW")	
IF diversion_check = 1 THEN Call write_variable_in_case_note ("* Diversion waiver")	
IF conversion_check = 1 THEN Call write_variable_in_case_note ("* Conversion waiver")	
IF essential_community_supports_check = 1 THEN Call write_variable_in_case_note ("* Essential Community supports.  Client does not meet LOC requirements.")	

'Information from DHS 5181 Dialog 2
Call write_bullet_and_variable_in_case_note ("Waiver Assessment Date", waiver_assessment_date_editbox)	
IF needs_waiver_check = 1 THEN Call write_variable_in_case_note ("* Waiver assessment date determined client needs waiver services & meets LOC requirements. Anticipated effective date no sooner than:", estimated_effective_date_editbox)		 
IF estimated_monthly_check  = 1 THEN Call write_variable_in_case_note ("* Estimated monthly waiver costs", estimated_monthly_waiver_costs_editbox)
IF does_not_meet_waiver_LOC_check = 1 THEN Call write_variable_in_case_note ("* Client does not meet LOC requirements for waivered services.")
IF ongoing_waiver_case_manager_check = 1 THEN Call write_variable_in_case_note ("* Ongoing case manager is", ongoing_waiver_case_manager_editbox)
Call write_bullet_and_variable_in_case_note ("LTCF Assessment Date", LTCF_assessment_date_editbox)	
IF meets_MALOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client meets the LOC requirement")
IF ongoing_case_manager_check = 1 THEN Call write_variable_in_case_note("* Ongoing case manager is", ongoing_case_manager_editbox)
IF ongoing_case_manager_not_available_check = 1 THEN Call write_variable_in_case_note ("* Ongoing Case Manager not available")
IF does_not_meet_MALTC_LOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client does not meet LOC requirements for LTCF's.")
IF requested_1503_check = 1 THEN Call write_variable_in_case_note ("* A DHS-1503 has been requested from the LTCF.")
IF onfile_1503_check = 1 THEN Call write_variable_in_case_note ("A DHS-1503 has been provided.")
IF client_applied_MA_check = 1 THEN Call write_variable_in_case_note ("* Client has applied for MA")
IF client_MA_enrollee_3543_provided_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client is an MA enrollee.  Assessor provided a DHS-3543 on:", Client_MA_enrollee_editbox)
IF completed_3543_3531_check = 1 THEN Call write_variable_in_case_note ("* Completed DHS-3543 or DHS-3531 attached to DHS 5181")
IF completed_3543_3531_faxed_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Completed DHS-3543 or DHS-3531 faxed to county on:", completed_3543_3531_faxed_editbox)
IF please_send_3543_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested that a DHS-3543 be sent to the MA enrollee or AREP.")
IF please_send_3531_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested that a DHS-3531 be sent to a non-MA enrollee at:", please_send_3531_check)
IF please_send_3340_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested an Asset Assessment, DHS 3340, be send to the client or AREP")

'Information from DHS 5181 Dialog 3
IF client_no_longer_meets_LOC_check = 1 THEN Call write_variable_in_case_note ("* Client no longer meets LOC - Effective date should be no sooner than:", client_no_longer_meets_LOC_efffective_date_editbox)					 
IF waiver_program_change_by_assessor_check = 1 THEN Call write_variable_in_case_note ("* Waiver program changed from:", waiver_program_change_from_assessor_editbox, "to", waiver_program_change_to_assessor_editbox)
Call write_bullet_and_variable_in_case_note ("Effective date", waiver_program_change_effective_date_editbox)
IF exited_waiver_program_check = 1 THEN Call write_variable_in_case_note("* Exited waiver program.  Effective date: ", exit_waiver_end_date_editbox)
IF client_choice_check = 1 THEN Call write_variable_in_case_note ("* Client has chosen to exit the waiver program")
IF client_deceased_check = 1 THEN Call write_variable_in_case_note ("* Client is deceased.  Date of death", date_of_death_editbox)
IF client_moved_to_LTCF_check = 1 THEN Call write_variable_in_case_note ("* Client moved to LTCF on", client_moved_to_LTCF_editbox)
Call write_bullet_and_variable_in_case_note ("Facility name", client_moved_to_LTCF_editbox)
Call write_bullet_and_variable_in_case_note ("Facility address", LTCF_ADDR_line_01 & LTCF_ADDR_line_02 & LTCF_city & LTCF_state & LTCF_zip_code)
IF waiver_program_change_check = 1 THEN Call write_variable_in_case_note ("* Waiver program changed from:", waiver_program_change_from_editbox, "to", waiver_program_change_to_editbox)
IF client_disenrolled_health_plan_check = 1 THEN Call write_variable_in_case_note ("* Client disenrolled from health plan effective", client_disenrolled_from_healthplan_editbox)
IF new_address_check = 1 THEN Call write_variable_in_case_note ("* New Address, effective date",new_address_effective_date_editbox & change_ADDR_line_1 & change_ADDR_line_2 & change_city & change_state & change_zip_code)
Call write_bullet_and_variable_in_case_note ("Case actions", case_action_editbox)
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes_editbox)
Call write_variable_in_case_note ("---")						 
call write_variable_in_case_note (worker_signature)
MsgBox "Make sure your DISA and FACI panel(s) are updated if they needed."

script_end_procedure("")



