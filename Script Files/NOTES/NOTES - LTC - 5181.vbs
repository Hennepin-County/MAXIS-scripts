OPTION EXPLICIT

'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - 5181.vbs"
start_time = timer

'FUNCTIONS LIBRARY
'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
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
DIM case_number_and_footer_month_dialog
DIM case_number_editbox
DIM month_editbox
DIM year_editbox
DIM next_month
DIM footer_month
DIM footer_year
DIM date_5181_editbox
DIM date_received_editbox
DIM lead_agency_assessor_editbox
DIM lead_agency_editbox
DIM address_editbox
DIM phone_editbox
DIM fax_editbox
DIM name_of_facility_editbox
DIM date_of_admission_editbox
DIM facility_address_editbox
DIM AC_check
DIM BI_check
DIM CAC_check
DIM CADI_check
DIM DD_check
DIM EW_check
DIM diversion_check
DIM conversion_check
DIM essential_community_supports_check
DIM waiver_assessment_date_editbox
DIM needs_waiver_checkbox
DIM estimated_effective_date_editbox
DIM estimated_monthly_check
DIM estimated_monthly_waiver_costs_editbox
DIM does_not_meet_waiver_LOC_check
DIM ongoing_case_manager_editbox
DIM LTCF_assessment_date_editbox
DIM meets_MALOC_check
DIM ongoing_case_manager_check
DIM ongoing_waiver_case_manager_check
DIM ongoing_waiver_case_manager_editbox
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
DIM next_button
DIM next_button_2
DIM previous_button
DIM client_no_longer_meets_LOC_check
DIM client_no_longer_meets_LOC_efffective_date_editbox
DIM waiver_program_change_check
DIM waiver_program_change_from_editbox
DIM waiver_program_change_to_editbox
DIM waiver_program_change_effective_date_editbox
DIM exited_waiver_program_check
DIM exit_waiver_end_date_editbox
DIM client_choice_check
DIM client_deceased_check
DIM date_of_death_editbox
DIM client_moved_to_LTCF_check
DIM client_moved_to_LTCF_editbox
DIM facility_name_edit
DIM waiver_program_change_by_assessor_check
DIM waiver_program_change_from_assessor_editbox
DIM waiver_program_change_to_assessor_editbox
DIM client_disenrolled_health_plan_check
DIM client_disenrolled_from_healthplan_editbox
DIM new_address_check
DIM new_address_editbox
DIM new_address_effective_date_editbox
DIM other_check
DIM other_changes_editbox
DIM Check14
DIM Previous_button_2
DIM previous_button_3
DIM write_TIKL_for_worker_check
DIM case_action_editbox
DIM other_notes_editbox
DIM sent_5181_to_caseworker_check
DIM requested_1503_check
DIM onfile_1503_check
DIM DHS_5181_dialog_1
DIM DHS_5181_dialog_2
DIM DHS_5181_dialog_3
DIM buttonpressed
DIM case_number
DIM case_note_dialog
DIM yes_case_note_button
DIM no_case_note_button
DIM yes_cancel_button
DIM cancel_dialog
DIM no_cancel_button
DIM maxis_footer_month
DIM maxis_footer_year
DIM row

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000

'FUNCTION----------------------------------------------------------------------------------------------------
FUNCTION cancel_confirmation 
	If ButtonPressed = 0 then  
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo) 
		If cancel_confirm = vbYes then stopscript 
	End if 
END FUNCTION 

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


BeginDialog DHS_5181_dialog_1, 0, 0, 361, 250, "5181 Dialog 1"
  EditBox 55, 5, 55, 15, date_5181_editbox
  EditBox 180, 5, 55, 15, date_received_editbox
  EditBox 135, 45, 95, 15, lead_agency_assessor_editbox
  EditBox 285, 45, 65, 15, lead_agency_editbox
  EditBox 110, 65, 235, 15, address_editbox
  EditBox 35, 85, 80, 15, phone_editbox
  EditBox 145, 85, 80, 15, fax_editbox
  EditBox 70, 135, 105, 15, name_of_facility_editbox
  EditBox 250, 135, 65, 15, date_of_admission_editbox
  EditBox 70, 155, 275, 15, facility_address_editbox
  CheckBox 10, 195, 20, 10, "AC", AC_check
  CheckBox 35, 195, 20, 10, "BI", BI_check
  CheckBox 60, 195, 30, 10, "CAC", CAC_check
  CheckBox 90, 195, 30, 10, "CADI", CADI_check
  CheckBox 125, 195, 25, 10, "DD", DD_check
  CheckBox 155, 195, 25, 10, "EW", EW_check
  CheckBox 250, 195, 40, 10, "Diversion", diversion_check
  CheckBox 300, 195, 50, 10, "Conversion", conversion_check
  CheckBox 40, 210, 190, 10, "Essential Community Supports (DHS- 3876 is required)", essential_community_supports_check
  ButtonGroup ButtonPressed
    CancelButton 305, 230, 50, 15
  Text 5, 5, 50, 15, "Date on 5181:"
  Text 125, 5, 55, 15, "Date Received:"
  Text 5, 65, 105, 15, "Address (include city/state/zip):"
  Text 125, 90, 20, 10, "Fax:"
  Text 5, 135, 60, 15, "Name of Facility:"
  Text 235, 50, 50, 10, "Lead Agency:"
  Text 185, 135, 65, 15, "Date of admission:"
  Text 5, 85, 30, 10, "Phone:"
  Text 5, 160, 60, 15, "Facility address:"
  Text 5, 45, 130, 10, "Lead Agency Assessor/Case Manager:"
  Text 195, 195, 45, 10, "Choose one:"
  Text 5, 30, 105, 15, "**CONTACT INFORMATION**"
  Text 10, 180, 285, 15, "OR The client is currently requesting services/enrolled in the following waiver program:"
  Text 5, 120, 45, 15, "**STATUS**"
  Text 15, 210, 15, 15, "OR:"
  GroupBox 0, 20, 355, 85, ""
  GroupBox 0, 110, 355, 115, ""
  ButtonGroup ButtonPressed
    PushButton 245, 230, 55, 15, "Next", next_button
EndDialog


BeginDialog DHS_5181_dialog_2, 0, 0, 361, 340, "5181 Dialog 2"
  EditBox 250, 20, 45, 15, waiver_assessment_date_editbox
  CheckBox 5, 35, 315, 10, "Needs waiver services and meets LOC requirement", needs_waiver_checkbox
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
  ButtonGroup ButtonPressed
    PushButton 250, 320, 50, 15, "Next", next_button_2
    CancelButton 305, 320, 50, 15
  Text 5, 20, 165, 15, "**WAIVERS** Assessment date determine client:"
  Text 145, 115, 60, 15, "Assessment date:"
  Text 20, 50, 135, 15, "Anticipated effective date no sooner than:"
  Text 5, 195, 190, 15, "**MEDICAL ASSISTANCE REQUESTS/APPLICATIONS**"
  Text 5, 0, 145, 15, "INITIAL REQUESTS (check all that apply):"
  Text 5, 115, 135, 15, "**LTCF** Assessment determines client: "
  Text 185, 20, 60, 15, "Assessment date:"
  GroupBox 0, 15, 355, 95, ""
  GroupBox 0, 110, 355, 80, ""
  GroupBox 0, 190, 355, 125, ""
  ButtonGroup ButtonPressed
    PushButton 195, 320, 50, 15, "Previous", previous_button_2
  CheckBox 30, 325, 65, 10, "1503 requested?", requested_1503_check
  CheckBox 105, 325, 55, 10, "1503 on file?", onfile_1503_check
EndDialog


BeginDialog DHS_5181_Dialog_3, 0, 0, 361, 325, "5181 Dialog 3"
  CheckBox 0, 20, 235, 15, "Client no longer meets LOC - Effective date should be no sooner than: ", client_no_longer_meets_LOC_check
  EditBox 240, 20, 45, 15, client_no_longer_meets_LOC_efffective_date_editbox
  CheckBox 0, 40, 90, 10, "Waiver program change", waiver_program_change_by_assessor_check
  EditBox 115, 40, 45, 15, waiver_program_change_from_assessor_editbox
  EditBox 180, 40, 45, 15, waiver_program_change_to_assessor_editbox
  EditBox 60, 55, 45, 15, waiver_program_change_effective_date_editbox
  CheckBox 0, 100, 140, 10, "Exited waiver program.  Effective date: ", exited_waiver_program_check
  EditBox 145, 95, 45, 15, exit_waiver_end_date_editbox
  CheckBox 25, 110, 65, 15, "Client's choice", client_choice_check
  CheckBox 25, 125, 115, 15, "Client deceased.  Date of death:", client_deceased_check
  EditBox 155, 125, 45, 15, date_of_death_editbox
  CheckBox 25, 140, 95, 15, "Client moved to LTCF on:", client_moved_to_LTCF_check
  EditBox 155, 140, 45, 15, client_moved_to_LTCF_editbox
  Text 55, 160, 50, 15, "Facility name:"
  EditBox 110, 160, 120, 15, facility_name_edit
  CheckBox 0, 180, 90, 15, "Waiver program change", waiver_program_change_check
  EditBox 115, 180, 45, 15, waiver_program_change_from_editbox
  EditBox 180, 180, 45, 15, waiver_program_change_to_editbox
  CheckBox 0, 195, 175, 15, "Client disenrolled from health plan.  Effective date: ", client_disenrolled_health_plan_check
  EditBox 180, 195, 45, 15, client_disenrolled_from_healthplan_editbox
  CheckBox 0, 210, 55, 15, "New address:", new_address_check
  EditBox 60, 210, 280, 15, new_address_editbox
  EditBox 60, 225, 45, 15, new_address_effective_date_editbox
  ButtonGroup ButtonPressed
    OkButton 245, 305, 50, 15
    CancelButton 300, 305, 50, 15
  Text 5, 5, 165, 15, "**CHANGES COMPLETED BY THE ASSESSOR**"
  Text 165, 180, 15, 15, "To: "
  Text 5, 85, 125, 15, "**CHANGES** (check all that apply):"
  CheckBox 50, 395, 145, 15, "Check4", Check14
  Text 95, 180, 20, 15, "From:"
  Text 165, 40, 15, 15, "To: "
  Text 10, 55, 50, 15, "Effective Date:"
  Text 10, 225, 50, 15, "Effective Date:"
  Text 95, 40, 20, 15, "From:"
  ButtonGroup ButtonPressed
    PushButton 190, 305, 50, 15, "Previous", previous_button_3
  CheckBox 5, 295, 120, 10, "Inform worker of 5181 via TIKL?", write_TIKL_for_worker_check
  Text 5, 255, 45, 15, "Case Action:"
  EditBox 55, 255, 285, 15, case_action_editbox
  GroupBox 0, 0, 350, 75, ""
  GroupBox 0, 80, 350, 165, ""
  Text 5, 275, 45, 15, "Other notes:"
  EditBox 55, 275, 285, 15, other_notes_editbox
  CheckBox 5, 310, 170, 10, "Sent 5181 back to Case Manager?", sent_5181_to_caseworker_check
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

'Checking to see that we're in MAXIS
call check_for_MAXIS(True)

'function to navigate user to case note
Call navigate_to_screen ("case", "note")						
PF9	

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


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Information from DHS 5181 Dialog 1
Call write_variable_in_case_note ("~~~DHS 5181 rec'd~~~")											
Call write_bullet_and_variable_in_case_note ("Date of 5181", date_5181_editbox )		
Call write_bullet_and_variable_in_case_note ("Date received", date_received_editbox)			
Call write_bullet_and_variable_in_case_note ("Lead Agency Assessor/Case Manager:",lead_agency_assessor_editbox)		  
Call write_bullet_and_variable_in_case_note ("Lead Agency", lead_agency_editbox)				 
Call write_bullet_and_variable_in_case_note ("Address", address_editbox)					 
Call write_bullet_and_variable_in_case_note ("Phone", phone_editbox)	
Call write_bullet_and_variable_in_case_note ("Fax", fax_editbox)
Call write_bullet_and_variable_in_case_note ("Name of Facilty", name_of_facility_editbox)											
Call write_bullet_and_variable_in_case_note ("Date of admission", date_of_admission_editbox )		
Call write_bullet_and_variable_in_case_note ("Facility address", facility_address_editbox)	
IF AC_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--AC")		
IF BI_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--BI")	
IF CAC_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--CAC")	
IF CADI_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--CADI")	
IF DD_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--DD")	
IF EW_check = 1 THEN Call write_bullet_and_variable_in_case_note ("The client is currently requesting services/enrolled in the following waiver program--EW")	
IF diversion_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Diversion waiver")	
IF conversion_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Conversion waiver")	
IF essential_community_supports_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Essential Community supports.  Client does not meet LOC requirements.")	
'Information from DHS 5181 Dialog 2
Call write_bullet_and_variable_in_case_note (" Waiver Assessment Date", waiver_assessment_date_editbox)	
IF needs_waiver_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note ("Waiver assessment date determined client needs waiver services and meets LOC requirements.  Anticipated effect date no sooner than", estimated_effective_date_editbox )		 
IF estimated_monthly_check  = 1 THEN Call write_bullet_and_variable_in_case_note ("Estimated monthly waiver costs",estimated_monthly_waiver_costs_editbox)
IF needs_waiver_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note
IF does_not_meet_waiver_LOC_check = 1 THEN Call write_variable_in_case_note ("* Client does not meet LOC requirements for waivered services.")
IF ongoing_waiver_case_manager_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Ongoing case manager is", ongoing_waiver_case_manager_editbox)
Call write_bullet_and_variable_in_case_note ("LTCF Assessment Date", LTCF_assessment_date_editbox)	
IF meets_MALOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client meets the LOC requirement")
IF ongoing_case_manager_check = 1 THEN Call write_bullet_and_variable_in_case_note("Ongoing case manager is",ongoing_case_manager_editbox)
IF ongoing_case_manager_not_available_check = 1 THEN Call write_variable_in_case_note ("* Ongoing Case Manager not available")
IF does_not_meet_MALTC_LOC_check = 1 THEN Call write_variable_in_case_note ("* LTCF Assessment determines that client does not meet LOC requirements for LTCF's.")
IF requested_1503_check = 1 THEN Call write_variable_in_case_note ("* A DHS-1503 has been requested from the LTCF.")
IF onfile_1503_check = 1 THEN Call write_variable_in_case_note ("A DHS-1503 has been provided.")
IF client_applied_MA_check = 1 THEN Call write_variable_in_case_note ("* Client has applied for MA")
IF client_MA_enrollee_3543_provided_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client is an MA enrollee.  Assessor provided a DHS-3543 on:", Client_MA_enrollee_editbox)
IF completed_3543_3531_check = 1 THEN Call write_variable_in_case_note ("* Completed DHS-3543 or DHS-3531 attached to DHS 5181")
IF completed_3543_3531_faxed_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Completed DHS-3543 or DHS-3531 faxed to county on:", completed_3543_3531_faxed_editbox)
IF please_send_3543_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested that a DHS-3543 be sent to the MA enrollee or AREP.")
IF please_send_3531_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Case manager has requested that a DHS-3531 be sent to a non-MA enrollee at:", please_send_3531_check)
IF please_send_3340_check = 1 THEN Call write_variable_in_case_note ("* Case manager has requested an Asset Assessment, DHS 3340, be send to the client or AREP")
'Information from DHS 5181 Dialog 3
IF client_no_longer_meets_LOC_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client no longer meets LOC - Effective date should be no sooner than:", client_no_longer_meets_LOC_efffective_date_editbox)					 
IF waiver_program_change_by_assessor_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Waiver program changed from:", waiver_program_change_from_assessor_editbox, "to", waiver_program_change_to_assessor_editbox)
Call write_bullet_and_variable_in_case_note ("Effective date", waiver_program_change_effective_date_editbox)
IF exited_waiver_program_check = 1 THEN Call write_bullet_and_variable_in_case_note("Exited waiver program.  Effective date: ", exit_waiver_end_date_editbox)
IF client_choice_check = 1 THEN Call write_variable_in_case_note ("* Client has chosen to exit the waiver program")
IF client_deceased_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client is deceased.  Date of death", date_of_death_editbox)
IF client_moved_to_LTCF_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client moved to LTCF on", client_moved_to_LTCF_editbox)
Call write_bullet_and_variable_in_case_note ("Facility name", client_moved_to_LTCF_editbox)
IF waiver_program_change_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Waiver program changed from:", waiver_program_change_from_editbox, "to", waiver_program_change_to_editbox)
IF client_disenrolled_health_plan_check = 1 THEN Call write_bullet_and_variable_in_case_note ("Client disenrolled from health plan effective", client_disenrolled_from_healthplan_editbox)
IF new_address_check = 1 THEN Call write_bullet_and_variable_in_case_note ("New Address",new_address_effective_date_editbox)
Call write_bullet_and_variable_in_case_note ("Effective date", new_address_effective_date_editbox)




IF subsidized_amount_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note ("Subsidized amount", subsidized_amount_editbox)			   
IF garage_amount_checkbox = 1 THEN Call write_bullet_and_variable_in_case_note ("Garage amount", garage_amount_editbox)				  
Call write_bullet_and_variable_in_case_note ("Utilities paid by resident", utilities_paid_by_resident_listbox) 
Call write_bullet_and_variable_in_case_note ("Other notes", other_notes_editbox)				
IF signed_by_LLMgr_checkbox = 1 THEN Call write_variable_in_case_note ("* Signed by LL/Mgr.")			  
IF signed_by_client_checkbox = 1 THEN Call write_variable_in_case_note ("* Signed by client.")
Call write_variable_in_case_note ("---")						 
call write_variable_in_case_note (worker_signature)

script_end_procedure ("")		

	