OPTION EXPLICIT

'STATS GATHERING ----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - 5181.vbs"
start_time = timer

'FUNCTIONS LIBRARY
LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
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
DIM ongoing_case_manager_check
DIM ongoing_case_manager_editbox
DIM LTCF_assessment_date_editbox
DIM meets_MALOC_check
DIM ongoing_case_manager_check
DIM ongoing_case_manager_editbox
DIM ongoing_case_manager_not_available_check
DIM does_not_meet_MALTC_LOC_check
DIM client_applied_MA_check
DIM client_MA_enrollee_3543_provided_check
DIM Client_MA_enrollee_editbox
DIM completed_3543_3531_ check
DIM completed_3543_3531_faxed_check
DIM completed_3543_3531_faxed_editbox
DIM please_send_3543_check
DIM please_send_3531_check
DIM please_send_3531_editbox
DIM please_send_3340_check
DIM next_button
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
DIM waiver_program_change_check
DIM waiver_program_change_from_editbox
DIM waiver_program_change_to_editbox
DIM client_disenrolled_health_plan_check
DIM client_disenrolled_from_healthplan_editbox
DIM new_address_check
DIM new_address_editbox
DIM new_address_effective_date_editbox
DIM other_check
DIM other_changes_editbox
DIM Check14
DIM previous_button
DIM write_TIKL_for_worker_check
DIM case_action_editbox
DIM other_notes_editbox
DIM sent_5181_to_caseworker_check

'DIALOGS
BeginDialog DHS_5181_dialog, 0, 0, 361, 250, "DHS-5181 Dialog 1"
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
  CheckBox 40, 210, 190, 10, "Esstential Community Supports (DHS- 3876 is required)", essential_community_supports_check
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


BeginDialog Dialog1, 0, 0, 361, 340, "5181 Dialog 2"
  EditBox 250, 20, 45, 15, waiver_assessment_date_editbox
  CheckBox 5, 35, 315, 10, "Needs waiver services and meets LOC requirement", needs_waiver_checkbox
  EditBox 160, 50, 45, 15, estimated_effective_date_editbox
  CheckBox 5, 65, 115, 15, "Estimated monthly waiver costs:", estimated_monthly_check
  EditBox 125, 65, 45, 15, estimated_monthly_waiver_costs_editbox
  CheckBox 5, 80, 170, 15, "Does not meet waiver services LOC requirement", does_not_meet_waiver_LOC_check
  CheckBox 5, 95, 105, 15, "Ongoing case mgr assigned:", ongoing_case_manager_check
  EditBox 110, 95, 60, 15, ongoing_case_manager_editbox
  EditBox 210, 115, 45, 15, LTCF_assessment_date_editbox
  CheckBox 5, 130, 100, 15, "Meets MA-LOC requirement", meets_MALOC_check
  CheckBox 5, 145, 120, 15, "Ongoing case manager assigned:", ongoing_case_manager_check
  EditBox 130, 145, 110, 15, ongoing_case_manager_editbox
  CheckBox 5, 160, 135, 15, "Ongoing case manager not available", ongoing_case_manager_not_available_check
  CheckBox 5, 175, 115, 15, "Does not meet LOC requirement", does_not_meet_MALTC_LOC_check
  CheckBox 0, 210, 80, 15, "Client applied for MA", client_applied_MA_check
  CheckBox 0, 225, 205, 15, "Client is an MA enrollee - Assessor provided DHS-3543 on:", client_MA_enrollee_3543_provided_check
  EditBox 205, 230, 45, 15, Client_MA_enrollee_editbox
  CheckBox 0, 240, 155, 15, "Completed DHS-3543 or DHS-3531 attached", completed_3543_3531_ check
  CheckBox 0, 255, 190, 15, "Completed DHS-3543 or DHS-3531 faxed to county on: ", completed_3543_3531_faxed_check
  EditBox 190, 255, 45, 15, completed_3543_3531_faxed_editbox
  CheckBox 0, 270, 180, 15, "Please send DHS-3543 to client (MA enrollee)", please_send_3543_check
  CheckBox 0, 285, 185, 15, "Please send DHS-3531 to client (Not MA enrollee) at:", please_send_3531_check
  EditBox 190, 285, 150, 15, please_send_3531_editbox
  CheckBox 0, 300, 205, 15, "Please send DHS-3340 to client - Asset Assessment needed", please_send_3340_check
  ButtonGroup ButtonPressed
    PushButton 250, 320, 50, 15, "Next", next_button
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
    PushButton 195, 320, 50, 15, "Previous", previous_button
EndDialog


BeginDialog 5181_Dialog_3, 0, 0, 361, 360, "5181 Dialog 3"
  CheckBox 0, 20, 235, 15, "Client no longer meets LOC - Effective date should be no sooner than: ", client_no_longer_meets_LOC_check
  EditBox 240, 20, 45, 15, client_no_longer_meets_LOC_efffective_date_editbox
  CheckBox 0, 40, 90, 10, "Waiver program change", waiver_program_change_check
  EditBox 115, 40, 45, 15, waiver_program_change_from_editbox
  EditBox 180, 40, 45, 15, waiver_program_change_to_editbox
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
  EditBox 60, 210, 290, 15, new_address_editbox
  EditBox 60, 225, 45, 15, new_address_effective_date_editbox
  CheckBox 45, 255, 30, 15, "Other:", other_check
  EditBox 35, 240, 315, 15, other_changes_editbox
  ButtonGroup ButtonPressed
    OkButton 245, 320, 50, 15
    CancelButton 300, 320, 50, 15
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
    PushButton 190, 320, 50, 15, "Previous", previous_button
  CheckBox 5, 310, 120, 10, "Inform worker of 5181 via TIKL?", write_TIKL_for_worker_check
  Text 5, 270, 45, 15, "Case Action:"
  EditBox 55, 270, 295, 15, case_action_editbox
  GroupBox 0, 0, 350, 75, ""
  GroupBox 0, 80, 350, 180, ""
  Text 5, 290, 45, 15, "Other notes:"
  EditBox 55, 290, 295, 15, other_notes_editbox
  CheckBox 5, 325, 170, 10, "Sent 5181 back to Case Manager?", sent_5181_to_caseworker_check
EndDialog

