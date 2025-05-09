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

'DEFINING CONSTANTS, ARRAY and BUTTONS===========================================================================

'Buttons Defined
'--Navigation buttons
section_a_contact_info_btn              = 201
section_b_contact_info_btn              = 202
status_btn                              = 203
initial_assessment_btn                  = 204
MA_req_app_btn                          = 205
exit_reasons_btn                        = 206
other_changes_btn                       = 207
section_d_comments_btn                  = 208
MA_status_determination_btn             = 209
changes_btn                             = 210  
section_g_comments_btn                  = 211
next_btn                                = 212
previous_btn                            = 213
complete_btn                            = 214

'--Other buttons
instructions_btn                        = 215
section_a_add_assessor_btn              = 216
section_e_add_assessor_btn              = 217
section_a_assessor_return_btn           = 218
section_e_assessor_return_btn           = 219
section_a_assessor_return_no_save_btn   = 220
section_e_assessor_return_no_save_btn   = 221
section_a_fill_SWKR_btn                 = 222
section_e_fill_SWKR_btn                 = 223
update_panels_btn                       = 224
skip_panel_updates_btn                  = 225

'Defining variables
dialog_count = ""
section_a_contact_info_called = False
section_a_additional_assessors_called = False
section_b_assess_results_current_status_called = False
section_b_assess_results_MA_requests_apps_changes_called = False
section_c_comm_elig_worker_exit_reasons_called = False
section_c_other_changes_section_d_comments_called = False
section_e_contact_info_called = False
section_e_additional_assessors_called = False
section_f_medical_assistance_called = False
section_f_medical_assistance_changes_called = False
section_g_comments_elig_worker_called = False

'DEFINING FUNCTIONS===========================================================================

'Dialog 1 - Section A: Contact Information
function section_a_contact_info()
  dialog_count = 1
  section_a_contact_info_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "1 - Section A: Contact Information"
    GroupBox 5, 5, 250, 190, "FROM (assessor/case manager/care coordinator's information)"
    Text 10, 20, 155, 10, "Click here to fill information from SWKR Panel:"
    ButtonGroup ButtonPressed
      PushButton 165, 15, 75, 15, "Fill from SWKR", section_a_fill_SWKR_btn
    Text 10, 40, 70, 10, "Date Sent to Worker:"
    EditBox 90, 35, 55, 15, section_a_date_form_sent
    Text 10, 55, 40, 10, "Assessor:"
    EditBox 90, 50, 150, 15, section_a_assessor
    Text 10, 70, 50, 10, "Lead Agency:"
    EditBox 90, 65, 150, 15, section_a_lead_agency
    Text 10, 85, 55, 10, "Phone Number:"
    EditBox 90, 80, 55, 15, section_a_phone_number
    Text 10, 100, 55, 10, "Street Address:"
    EditBox 90, 95, 150, 15, section_a_street_address
    Text 10, 115, 20, 10, "City:"
    EditBox 90, 110, 150, 15, section_a_city
    Text 10, 130, 25, 10, "State:"
    EditBox 90, 125, 25, 15, section_a_state
    Text 10, 145, 35, 10, "Zip Code:"
    EditBox 90, 140, 55, 15, section_a_zip_code
    Text 10, 160, 55, 10, "Email Address:"
    EditBox 90, 155, 150, 15, section_a_email_address
    Text 10, 180, 145, 10, "Click button to add up to 2 add'l assessors:"
    ButtonGroup ButtonPressed
      PushButton 160, 175, 85, 15, "Add/Update Assessor", section_a_add_assessor_btn
    GroupBox 5, 205, 250, 30, "Person's Information"
    Text 10, 215, 70, 10, "Select HH Member:"
    DropListBox 80, 215, 160, 15, HH_Memb_DropDown, hh_memb
    ButtonGroup ButtonPressed
      PushButton 215, 290, 55, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 55, 15, "Previous", previous_btn
    GroupBox 260, 5, 60, 280, "Navigation"
    Text 265, 15, 40, 10, "Section A"
    Text 265, 45, 40, 10, "Section B"
    Text 265, 105, 40, 10, "Section C"
    Text 265, 150, 40, 10, "Section D"
    Text 265, 180, 40, 10, "Section E"
    Text 265, 210, 40, 10, "Section F"
    Text 265, 255, 40, 10, "Section G"
    ButtonGroup ButtonPressed
      PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
'To do - dim all the variables?
Dim section_a_date_form_sent, section_a_assessor, section_a_lead_agency, section_a_phone_number, section_a_street_address, section_a_city, section_a_state, section_a_zip_code, section_a_email_address, hh_memb

function section_a_additional_assessors()
  dialog_count = 11
  section_a_additional_assessors_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "Section A: Contact Info (Add'l Assessors)"
    GroupBox 5, 5, 245, 135, "Additional Assessor (2)"
    Text 10, 20, 40, 10, "Assessor:"
    EditBox 90, 15, 150, 15, section_a_assessor_2
    Text 10, 35, 50, 10, "Lead Agency:"
    EditBox 90, 30, 150, 15, section_a_lead_agency_2
    Text 10, 50, 55, 10, "Phone Number:"
    EditBox 90, 45, 55, 15, section_a_phone_number_2
    Text 10, 65, 55, 10, "Street Address:"
    EditBox 90, 60, 150, 15, section_a_street_address_2
    Text 10, 80, 20, 10, "City:"
    EditBox 90, 75, 150, 15, section_a_city_2
    Text 10, 95, 25, 10, "State:"
    EditBox 90, 90, 25, 15, section_a_state_2
    Text 10, 110, 35, 10, "Zip Code:"
    EditBox 90, 105, 55, 15, section_a_zip_code_2
    Text 10, 125, 55, 10, "Email Address:"
    EditBox 90, 120, 150, 15, section_a_email_address_2
    GroupBox 5, 150, 245, 135, "Additional Assessor (3) - Leave all fields blank if unneeded"
    Text 10, 165, 40, 10, "Assessor:"
    EditBox 90, 160, 150, 15, section_a_assessor_3
    Text 10, 180, 50, 10, "Lead Agency:"
    EditBox 90, 175, 150, 15, section_a_lead_agency_3
    Text 10, 195, 55, 10, "Phone Number:"
    EditBox 90, 190, 55, 15, section_a_phone_number_3
    Text 10, 210, 55, 10, "Street Address:"
    EditBox 90, 205, 150, 15, section_a_street_address_3
    Text 10, 225, 20, 10, "City:"
    EditBox 90, 220, 150, 15, section_a_city_3
    Text 10, 240, 25, 10, "State:"
    EditBox 90, 235, 25, 15, section_a_state_3
    Text 10, 255, 35, 10, "Zip Code:"
    EditBox 90, 250, 55, 15, section_a_zip_code_3
    Text 10, 270, 55, 10, "Email Address:"
    EditBox 90, 265, 150, 15, section_a_email_address_3
    ButtonGroup ButtonPressed
      PushButton 195, 290, 130, 15, "Save Info and Return to Contact Info", section_a_assessor_return_btn
      PushButton 5, 290, 185, 15, "Return to Contact Info WITHOUT Saving Assessor Info", section_a_assessor_return_no_save_btn
  EndDialog
end function
'Dim all variables in function
Dim section_a_assessor_2, section_a_lead_agency_2, section_a_phone_number_2, section_a_street_address_2, section_a_city_2, section_a_state_2, section_a_zip_code_2, section_a_email_address_2, section_a_assessor_3, section_a_lead_agency_3, section_a_phone_number_3, section_a_street_address_3, section_a_city_3, section_a_state_3, section_a_zip_code_3, section_a_email_address_3

'Dialog 2 - Section B: Assessment Results - Current Status
function section_b_assess_results_current_status()
  dialog_count = 2
  section_b_assess_results_current_status_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "2 - Section B: Assess. Results - Current Status"
  GroupBox 5, 5, 250, 50, "What is the person's current status? (check second if both apply)"
  CheckBox 15, 20, 10, 10, "", section_g_person_requesting_already_enrolled_LTC
  Text 25, 20, 215, 20, "The person currently is requesting services or already enrolled in long-term care services or program"
  CheckBox 15, 40, 195, 10, "The person resides in or will reside in an institution", section_g_person_will_reside_institution_checkbox
  GroupBox 5, 60, 250, 55, "Program Type:"
  Text 10, 75, 185, 10, "Program person is requesting or is currently enrolled in:"
  DropListBox 195, 70, 55, 20, "Select one:"+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EQ"+chr(9)+"ECS"+chr(9)+"PCA/CFSS", section_b_program_type
  Text 10, 90, 85, 10, "Check one (if applicable):"
  CheckBox 105, 85, 45, 15, "Diversion", section_b_diversion_checkbox
  CheckBox 155, 85, 50, 15, "Conversion", section_b_conversion_checkbox
  GroupBox 5, 120, 245, 125, "Institution"
  Text 15, 135, 60, 10, "Admission Date:"
  Text 15, 150, 60, 10, "Facility:"
  Text 15, 165, 60, 10, "Phone Number:"
  Text 15, 180, 60, 10, "Street Address:"
  Text 15, 195, 60, 10, "City:"
  Text 15, 210, 60, 10, "State:"
  Text 15, 225, 60, 10, "Zip Code:"
  EditBox 80, 130, 95, 15, section_b_admission_date
  EditBox 80, 145, 95, 15, section_b_facility
  EditBox 80, 160, 95, 15, section_b_institution_phone_number
  EditBox 80, 175, 95, 15, section_b_institution_street_address
  EditBox 80, 190, 95, 15, section_b_institution_city
  EditBox 80, 205, 95, 15, section_b_institution_state
  EditBox 80, 220, 95, 15, section_b_institution_zip_code
  ButtonGroup ButtonPressed
    PushButton 215, 290, 55, 15, "Next", next_btn
    CancelButton 270, 290, 50, 15
    PushButton 5, 290, 55, 15, "Previous", previous_btn
  GroupBox 260, 5, 60, 280, "Navigation"
  Text 265, 15, 40, 10, "Section A"
  Text 265, 45, 40, 10, "Section B"
  Text 265, 105, 40, 10, "Section C"
  Text 265, 150, 40, 10, "Section D"
  Text 265, 180, 40, 10, "Section E"
  Text 265, 210, 40, 10, "Section F"
  Text 265, 255, 40, 10, "Section G"
  ButtonGroup ButtonPressed
    PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
'Dim all variables in function
Dim section_g_person_requesting_already_enrolled_LTC, section_g_person_will_reside_institution_checkbox, section_b_program_type, section_b_diversion_checkbox, section_b_conversion_checkbox, section_b_admission_date, section_b_facility, section_b_institution_phone_number, section_b_institution_street_address, section_b_institution_city, section_b_institution_state, section_b_institution_zip_code

'Dialog 3 - Section B: Assessment Results - Initial Assessment & Case Manager
function section_b_assess_results_initial_assess_case_manager()
  dialog_count = 3
  section_b_assess_results_initial_assess_case_manager_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "3 - Section B: Assess. Results - Initial Assess. and Case Manager"
  GroupBox 5, 5, 240, 105, "Initial Assessment"
  Text 15, 20, 55, 10, "Assessment on "
  EditBox 70, 15, 55, 15, section_b_assessment_date
  Text 130, 20, 90, 10, "determined that the person:"
  DropListBox 15, 35, 200, 20, "Select one:"+chr(9)+"Does not meet institutional LOC requirement"+chr(9)+"Meets institutional LOC requirement"+chr(9)+"Meets need criteria for PCA/CFSS", section_b_assessment_determination
  Text 15, 60, 135, 10, "Will the person open to waiver/AC/ECS?"
  CheckBox 155, 55, 25, 15, "Yes", section_b_open_to_waiver_yes_checkbox
  CheckBox 185, 55, 25, 15, "No", section_b_open_to_waiver_no_checkbox
  Text 15, 80, 120, 10, "Estimated monthly waiver/AC costs:"
  EditBox 140, 75, 70, 15, section_b_monthly_waiver_costs
  Text 15, 95, 90, 10, "Anticipated effective date:"
  EditBox 140, 90, 70, 15, section_b_waiver_effective_date
  GroupBox 5, 120, 240, 115, "Case Manager"
  Text 15, 135, 220, 10, "Does the person have a case manager? (select ONE option below)"
  CheckBox 20, 150, 105, 10, "Yes - I am the case manager", section_b_yes_case_manager
  CheckBox 20, 165, 130, 10, "Yes - Someone else is the manager", section_b_yes_someone_else_case_manager
  CheckBox 20, 180, 125, 10, "No (enter case manager info below)", section_b_no_case_manager
  Text 35, 200, 75, 10, "Case Manager Name:"
  EditBox 110, 195, 130, 15, section_b_case_manager_name
  Text 35, 215, 60, 10, "Phone Number:"
  EditBox 110, 210, 70, 15, section_b_case_manager_phone_number
  ButtonGroup ButtonPressed
    PushButton 215, 290, 55, 15, "Next", next_btn
    CancelButton 270, 290, 50, 15
    PushButton 5, 290, 55, 15, "Previous", previous_btn
  GroupBox 260, 5, 60, 280, "Navigation"
  Text 265, 15, 40, 10, "Section A"
  Text 265, 45, 40, 10, "Section B"
  Text 265, 105, 40, 10, "Section C"
  Text 265, 150, 40, 10, "Section D"
  Text 265, 180, 40, 10, "Section E"
  Text 265, 210, 40, 10, "Section F"
  Text 265, 255, 40, 10, "Section G"
  ButtonGroup ButtonPressed
    PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
'Dim all variables in function
Dim section_b_assessment_date, section_b_assessment_determination, section_b_open_to_waiver_yes_checkbox, section_b_open_to_waiver_no_checkbox, section_b_monthly_waiver_costs, section_b_waiver_effective_date, section_b_yes_case_manager, section_b_yes_someone_else_case_manager, section_b_no_case_manager, section_b_case_manager_name, section_b_case_manager_phone_number


'Dialog 4 - Section B: Assessment Results - MA Requests/Apps & Changes
function section_b_assess_results_MA_requests_apps_changes()
  dialog_count = 4
  section_b_assess_results_MA_requests_apps_changes_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "4 - Section B: Assess. Results - MA Requests/Apps and Changes"
    GroupBox 5, 5, 250, 200, "Medical Assistance requests/applications (select all that apply):"
    CheckBox 15, 15, 110, 10, "Person applied for MA/MA-LTC", section_b_applied_MA_LTC_checkbox
    CheckBox 15, 30, 110, 10, "Person is an MA enrollee", section_b_ma_enrollee_checkbox
    Text 30, 40, 170, 10, "What date did the assessor provide the DHS-3543?"
    EditBox 205, 35, 40, 15, section_b_date_dhs_3543_provided
    CheckBox 15, 55, 205, 10, "Person completed DHS-3543 or DHS-3531 and it is attached", section_b_completed_dhs_3543_3531_attached_checkbox
    CheckBox 15, 70, 150, 10, "Person completed DHS-3543 or DHS-3531", section_b_completed_dhs_3543_3531_checkbox
    Text 30, 80, 80, 10, "Date sent to the county:"
    EditBox 110, 80, 40, 15, section_b_dhs_3543_3531_sent_to_county_date
    CheckBox 15, 100, 145, 10, "Send DHS-3543 to person (MA Enrollee)", section_b_send_dhs_3543_checkbox
    CheckBox 15, 115, 160, 10, "Send DHS-3531 to person (Not MA Enrollee)", section_b_send_dhs_3531_checkbox
    EditBox 25, 125, 80, 15, section_b_send_dhs_3531_address
    EditBox 105, 125, 60, 15, section_b_send_dhs_3531_city
    EditBox 165, 125, 25, 15, section_b_send_dhs_3531_state
    EditBox 190, 125, 40, 15, section_b_send_dhs_3531_zip
    Text 50, 140, 175, 10, "Address                          City              State        Zip"
    CheckBox 15, 155, 190, 10, "Send DHS-3340 to person (asset assessment needed)", section_b_send_dhs_3340_checkbox
    EditBox 25, 170, 80, 15, section_b_send_dhs_3340_address
    EditBox 105, 170, 60, 15, section_b_send_dhs_3340_city
    EditBox 165, 170, 25, 15, section_b_send_dhs_3340_state
    EditBox 190, 170, 40, 15, section_b_send_dhs_3340_zip
    Text 50, 185, 175, 10, "Address                          City              State        Zip"
    GroupBox 5, 215, 250, 55, "Changes completed by assessor at reassessment (select all that apply)"
    CheckBox 15, 225, 145, 10, "Person no longer meets institutional LOC", section_b_person_no_longer_institutional_LOC_checkbox
    Text 30, 235, 170, 10, "Effect. date of waiver exit should be no sooner than:"
    EditBox 205, 230, 45, 15, section_b_date_waiver_exit
    CheckBox 15, 250, 155, 10, "Person chooses to enroll in another program", section_b_person_enroll_another_program
    DropListBox 180, 250, 70, 15, "Select one:"+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EQ"+chr(9)+"ECS"+chr(9)+"PCA/CFSS", section_b_enroll_another_program_list
    ButtonGroup ButtonPressed
      PushButton 215, 290, 55, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 55, 15, "Previous", previous_btn
    GroupBox 260, 5, 60, 280, "Navigation"
    Text 265, 15, 40, 10, "Section A"
    Text 265, 45, 40, 10, "Section B"
    Text 265, 105, 40, 10, "Section C"
    Text 265, 150, 40, 10, "Section D"
    Text 265, 180, 40, 10, "Section E"
    Text 265, 210, 40, 10, "Section F"
    Text 265, 255, 40, 10, "Section G"
    ButtonGroup ButtonPressed
      PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
'Dim all variables in function
Dim section_b_applied_MA_LTC_checkbox, section_b_ma_enrollee_checkbox, section_b_date_dhs_3543_provided, section_b_completed_dhs_3543_3531_attached_checkbox, section_b_completed_dhs_3543_3531_checkbox, section_b_dhs_3543_3531_sent_to_county_date, section_b_send_dhs_3543_checkbox, section_b_send_dhs_3531_checkbox, section_b_send_dhs_3531_address, section_b_send_dhs_3531_city, section_b_send_dhs_3531_state, section_b_send_dhs_3531_zip, section_b_send_dhs_3340_checkbox, section_b_send_dhs_3340_address, section_b_send_dhs_3340_city, section_b_send_dhs_3340_state, section_b_send_dhs_3340_zip, section_b_person_no_longer_institutional_LOC_checkbox, section_b_date_waiver_exit, section_b_person_enroll_another_program, section_b_enroll_another_program_list

'Dialog 5 - Section C: Communication to eligibility worker - Exit Reasons
function section_c_comm_elig_worker_exit_reasons()
  dialog_count = 5
  section_c_comm_elig_worker_exit_reasons_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "5 - Section C: Comm. to elig. worker - Exit Reasons"
    GroupBox 5, 5, 245, 200, "Exit Reasons"
    CheckBox 15, 20, 125, 10, "The person exited waiver program", section_c_exited_waiver_program_checkbox
    Text 35, 35, 95, 10, "Effective date of waiver exit:"
    EditBox 135, 30, 40, 15, section_c_date_waiver_exit
    Text 15, 55, 175, 10, "Reason - Check reason for exit (select all that apply)"
    CheckBox 20, 70, 75, 10, "Hospital admission", section_c_hospital_admission_checkbox
    CheckBox 20, 80, 95, 10, "Nursing facility admission", section_c_nursing_facility_admission_checkbox
    CheckBox 20, 100, 100, 10, "Person's informed choice", section_c_person_informed_choice_checkbox
    CheckBox 20, 90, 115, 10, "Residential treatment admission", section_c_residential_treatment_admission_checkbox
    CheckBox 20, 110, 85, 10, "Person is deceased", section_c_person_deceased_checkbox
    Text 35, 125, 50, 10, "Date of death:"
    EditBox 85, 120, 40, 15, section_c_date_of_death
    CheckBox 20, 135, 110, 10, "Person moved out of state", section_c_person_moved_out_of_state_checkbox
    Text 35, 150, 50, 10, "Date of move:"
    EditBox 85, 145, 40, 15, section_c_date_of_move
    CheckBox 20, 165, 205, 10, "Exited for other reasons (not including LOC) - explain below:", section_c_exited_for_other_reasons_checkbox
    EditBox 30, 180, 215, 15, section_c_exited_for_other_reasons_explanation
    ButtonGroup ButtonPressed
      PushButton 215, 290, 55, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 55, 15, "Previous", previous_btn
    GroupBox 260, 5, 60, 280, "Navigation"
    Text 265, 15, 40, 10, "Section A"
    Text 265, 45, 40, 10, "Section B"
    Text 265, 105, 40, 10, "Section C"
    Text 265, 150, 40, 10, "Section D"
    Text 265, 180, 40, 10, "Section E"
    Text 265, 210, 40, 10, "Section F"
    Text 265, 255, 40, 10, "Section G"
    ButtonGroup ButtonPressed
      PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
' Dim all variables in function
Dim section_c_exited_waiver_program_checkbox, section_c_date_waiver_exit, section_c_hospital_admission_checkbox, section_c_nursing_facility_admission_checkbox, section_c_person_informed_choice_checkbox, section_c_residential_treatment_admission_checkbox, section_c_person_deceased_checkbox, section_c_date_of_death, section_c_person_moved_out_of_state_checkbox, section_c_date_of_move, section_c_exited_for_other_reasons_checkbox, section_c_exited_for_other_reasons_explanation

'Dialog 6 - Section C: Other Changes & Section D: Comments
function section_c_other_changes_section_d_comments()
  dialog_count = 6
  section_c_other_changes_section_d_comments_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "6 - Section C: Other Changes & Section D: Comments"
  GroupBox 5, 5, 250, 235, "Other changes"
  Text 15, 20, 50, 10, "Program type:"
  DropListBox 70, 15, 65, 15, "Select one:"+chr(9)+"AC"+chr(9)+"BI"+chr(9)+"CAC"+chr(9)+"CADI"+chr(9)+"DD"+chr(9)+"EW"+chr(9)+"ECS"+chr(9)+"PCA/CFSS", section_c_program_type_list
  Text 20, 30, 90, 10, "Choose one (if applicable)"
  CheckBox 120, 30, 45, 10, "Diversion", section_c_diversion_checkbox
  CheckBox 170, 30, 50, 10, "Conversion", section_c_conversion_checkbox
  Text 15, 45, 100, 10, "Changes (select all that apply)"
  CheckBox 15, 60, 145, 10, "Person has moved or has a new address", section_c_person_moved_new_address_checkbox
  Text 35, 75, 80, 10, "Date address changed:"
  EditBox 120, 70, 30, 15, section_c_date_address_changed
  EditBox 25, 90, 70, 15, section_c_street_address
  EditBox 95, 90, 50, 15, section_c_city
  EditBox 145, 90, 25, 15, section_c_state
  EditBox 170, 90, 40, 15, section_c_zip_code
  Text 35, 105, 165, 10, "Address                       City                State       Zip"
  CheckBox 15, 115, 205, 10, "Person has a new legal representative (enter details below)", section_c_new_legal_rep_checkbox
  EditBox 25, 125, 80, 15, section_c_legal_rep_first_name
  EditBox 105, 125, 80, 15, section_c_legal_rep_last_name
  EditBox 185, 125, 55, 15, section_c_legal_rep_phone_number
  Text 35, 140, 195, 10, "First name                         Last name                      Phone number"
  EditBox 25, 150, 70, 15, section_c_legal_rep_street_address
  EditBox 95, 150, 50, 15, section_c_legal_rep_city
  EditBox 145, 150, 25, 15, section_c_legal_rep_state
  EditBox 170, 150, 40, 15, section_c_legal_rep_zip_code
  Text 35, 165, 165, 10, "Address                       City                State       Zip"
  CheckBox 15, 175, 225, 10, "Person returning to community w/in 121 days of a qual. admission", section_c_person_return_to_community_checkbox
  Text 40, 190, 50, 10, "Effective date:"
  EditBox 95, 185, 30, 15, section_c_qual_admission_eff_date
  CheckBox 15, 205, 225, 10, "Other changes related to program/service elig. (describe changes)", section_c_other_changes_program_checkbox
  EditBox 25, 220, 225, 15, section_c_other_changes_program
  GroupBox 5, 245, 250, 40, "Section D: Comments from assessor, case manager or care coordinator"
  Text 15, 255, 215, 10, "Enter any additional notes or comments"
  EditBox 15, 265, 225, 15, section_d_additional_comments
  ButtonGroup ButtonPressed
    PushButton 215, 290, 55, 15, "Next", next_btn
    CancelButton 270, 290, 50, 15
    PushButton 5, 290, 55, 15, "Previous", previous_btn
  GroupBox 260, 5, 60, 280, "Navigation"
  Text 265, 15, 40, 10, "Section A"
  Text 265, 45, 40, 10, "Section B"
  Text 265, 105, 40, 10, "Section C"
  Text 265, 150, 40, 10, "Section D"
  Text 265, 180, 40, 10, "Section E"
  Text 265, 210, 40, 10, "Section F"
  Text 265, 255, 40, 10, "Section G"
  ButtonGroup ButtonPressed
    PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
EndDialog
end function
'Dim the variables in function
Dim section_c_program_type_list, section_c_diversion_checkbox, section_c_conversion_checkbox, section_c_person_moved_new_address_checkbox, section_c_date_address_changed, section_c_street_address, section_c_city, section_c_state, section_c_zip_code, section_c_new_legal_rep_checkbox, section_c_legal_rep_first_name, section_c_legal_rep_last_name, section_c_legal_rep_phone_number, section_c_legal_rep_street_address, section_c_legal_rep_city, section_c_legal_rep_state, section_c_legal_rep_zip_code, section_c_person_return_to_community_checkbox, section_c_qual_admission_eff_date, section_c_other_changes_program_checkbox, section_c_other_changes_program, section_d_additional_comments

'Dialog 7 - Section E: Contact Information
function section_e_contact_info()
  first_name = replace(first_name, "_", "")
  last_name = replace(last_name, "_", "")
  ref_nbr = left(hh_memb, 2)
  dialog_count = 7
  section_e_contact_info_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "7 - Section E: Contact Information"
    Text 10, 25, 180, 10, "Date Sent to assessor/case manager/care coordinator:"
    EditBox 195, 20, 55, 15, section_e_date_form_sent
    GroupBox 5, 5, 250, 200, "TO (assessor/case manager/care coordinator's information)"
    Text 10, 45, 155, 10, "Click here to fill information from SWKR Panel:"
    ButtonGroup ButtonPressed
      PushButton 165, 40, 75, 15, "Fill from SWKR", section_e_fill_SWKR_btn
    Text 10, 65, 40, 10, "Assessor:"
    EditBox 90, 60, 150, 15, section_e_assessor
    Text 10, 80, 50, 10, "Lead Agency:"
    EditBox 90, 75, 150, 15, section_e_lead_agency
    Text 10, 95, 55, 10, "Phone Number:"
    EditBox 90, 90, 55, 15, section_e_phone_number
    Text 10, 110, 55, 10, "Street Address:"
    EditBox 90, 105, 150, 15, section_e_street_address
    Text 10, 125, 20, 10, "City:"
    EditBox 90, 120, 150, 15, section_e_city
    Text 10, 140, 25, 10, "State:"
    EditBox 90, 135, 25, 15, section_e_state
    Text 10, 155, 35, 10, "Zip Code:"
    EditBox 90, 150, 45, 15, section_e_zip_code
    Text 10, 170, 55, 10, "Email Address:"
    EditBox 90, 165, 150, 15, section_e_email_address
    Text 10, 190, 145, 10, "Click button to add up to 2 add'l assessors:"
    ButtonGroup ButtonPressed
      PushButton 160, 185, 85, 15, "Add/Update Assessor", section_e_add_assessor_btn
    GroupBox 5, 215, 245, 60, "Person's Information"
    Text 10, 230, 105, 10, "Information entered previously:"
    Text 15, 240, 40, 10, "First name:"
    Text 70, 240, 170, 10, first_name
    Text 15, 250, 40, 10, "Last name:"
    Text 70, 250, 170, 10, last_name
    Text 15, 260, 45, 10, "Ref Number:"
    Text 70, 260, 75, 10, ref_nbr
    ButtonGroup ButtonPressed
      PushButton 215, 290, 55, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 55, 15, "Previous", previous_btn
    GroupBox 260, 5, 60, 280, "Navigation"
    Text 265, 15, 40, 10, "Section A"
    Text 265, 45, 40, 10, "Section B"
    Text 265, 105, 40, 10, "Section C"
    Text 265, 150, 40, 10, "Section D"
    Text 265, 180, 40, 10, "Section E"
    Text 265, 210, 40, 10, "Section F"
    Text 265, 255, 40, 10, "Section G"
    ButtonGroup ButtonPressed
      PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
' Dim all variables in function
Dim section_e_date_form_sent, section_e_assessor, section_e_lead_agency, section_e_phone_number, section_e_street_address, section_e_city, section_e_state, section_e_zip_code, section_e_email_address, first_name, last_name, ref_nbr

'Dialog 7 - Section E: Contact Information
function section_e_additional_assessors()
  dialog_count = 12
  section_e_additional_assessors_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "Section E: Contact Info (Add'l Assessors)"
    GroupBox 5, 5, 245, 135, "Additional Assessor (2)"
    Text 10, 20, 40, 10, "Assessor:"
    EditBox 90, 15, 150, 15, section_e_assessor_2
    Text 10, 35, 50, 10, "Lead Agency:"
    EditBox 90, 30, 150, 15, section_e_lead_agency_2
    Text 10, 50, 55, 10, "Phone Number:"
    EditBox 90, 45, 55, 15, section_e_phone_number_2
    Text 10, 65, 55, 10, "Street Address:"
    EditBox 90, 60, 150, 15, section_e_street_address_2
    Text 10, 80, 20, 10, "City:"
    EditBox 90, 75, 150, 15, section_e_city_2
    Text 10, 95, 25, 10, "State:"
    EditBox 90, 90, 25, 15, section_e_state_2
    Text 10, 110, 35, 10, "Zip Code:"
    EditBox 90, 105, 55, 15, section_e_zip_code_2
    Text 10, 125, 55, 10, "Email Address:"
    EditBox 90, 120, 150, 15, section_e_email_address_2
    GroupBox 5, 150, 245, 135, "Additional Assessor (3) - Leave blank if unneeded"
    Text 10, 165, 40, 10, "Assessor:"
    EditBox 90, 160, 150, 15, section_e_assessor_3
    Text 10, 180, 50, 10, "Lead Agency:"
    EditBox 90, 175, 150, 15, section_e_lead_agency_3
    Text 10, 195, 55, 10, "Phone Number:"
    EditBox 90, 190, 55, 15, section_e_phone_number_3
    Text 10, 210, 55, 10, "Street Address:"
    EditBox 90, 205, 150, 15, section_e_street_address_3
    Text 10, 225, 20, 10, "City:"
    EditBox 90, 220, 150, 15, section_e_city_3
    Text 10, 240, 25, 10, "State:"
    EditBox 90, 235, 25, 15, section_e_state_3
    Text 10, 255, 35, 10, "Zip Code:"
    EditBox 90, 250, 55, 15, section_e_zip_code_3
    Text 10, 270, 55, 10, "Email Address:"
    EditBox 90, 265, 150, 15, section_e_email_address_3
    ButtonGroup ButtonPressed
      PushButton 195, 290, 130, 15, "Save Info and Return to Contact Info", section_e_assessor_return_btn
      PushButton 5, 290, 185, 15, "Return to Contact Info WITHOUT Saving Assessor Info", section_e_assessor_return_no_save_btn
  EndDialog
end function
' Dim all functions in variable
Dim section_e_assessor_2, section_e_lead_agency_2, section_e_phone_number_2, section_e_street_address_2, section_e_city_2, section_e_state_2, section_e_zip_code_2, section_e_email_address_2, section_e_assessor_3,  section_e_lead_agency_3, section_e_phone_number_3, section_e_street_address_3, section_e_city_3, section_e_state_3, section_e_zip_code_3, section_e_email_address_3

'Dialog 8 - Section F: Medical Assistance
function section_f_medical_assistance()
  dialog_count = 8
  section_f_medical_assistance_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "8 - Section F: Medical Assistance"
  CheckBox 15, 20, 175, 10, "Person applied for MA/MA-LTC (enter date applied)", section_f_person_applied_MA_checkbox
  CheckBox 15, 35, 150, 10, "DHS-3531 sent to person (enter date sent)", section_f_dhs_3531_sent_checkbox
  EditBox 195, 15, 30, 15, section_f_person_applied_date
  EditBox 195, 30, 30, 15, section_f_dhs_3531_sent_date
  CheckBox 15, 50, 150, 10, "DHS-3543 sent to person (enter date sent)", section_f_dhs_3543_sent_checkbox
  EditBox 195, 45, 30, 15, section_f_dhs_3543_sent_date
  CheckBox 15, 65, 200, 10, "DHS-3543/DHS-3531 returned; elig determination pending", section_f_dhs_3543_3531_returned_checkbox
  Text 30, 80, 40, 10, "Comments:"
  EditBox 70, 75, 160, 15, section_f_dhs_3543_3531_returned_comments
  CheckBox 15, 95, 155, 10, "DHS-3543/DHS-3531 has not been returned", section_f_dhs_3543_3531_not_returned_checkbox
  GroupBox 5, 5, 250, 105, "MA status for long-term supports and services (select all that apply)"
  ButtonGroup ButtonPressed
    CancelButton 270, 290, 50, 15
    PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  GroupBox 260, 5, 60, 280, "Navigation"
  Text 265, 15, 40, 10, "Section A"
  Text 265, 45, 40, 10, "Section B"
  Text 265, 105, 40, 10, "Section C"
  Text 265, 150, 40, 10, "Section D"
  Text 265, 180, 40, 10, "Section E"
  Text 265, 210, 40, 10, "Section F"
  Text 265, 255, 40, 10, "Section G"
  GroupBox 5, 115, 250, 160, "Determination (select all that apply)"
  CheckBox 10, 130, 100, 10, "MA opened (enter eff. date)", section_f_ma_opened_checkbox
  EditBox 210, 125, 30, 15, section_f_ma_opened_date
  CheckBox 10, 145, 165, 10, "Basic MA medical spenddown (enter amount $)", section_f_basic_ma_medical_spenddown_checkbox
  EditBox 210, 140, 30, 15, section_f_basic_ma_medical_spenddown
  CheckBox 10, 160, 200, 10, "MA for LTC services open on specific date (enter eff. date)", section_f_ma_LTC_services_checkbox
  EditBox 210, 155, 30, 15, section_f_ma_LTC_services_date
  CheckBox 10, 175, 195, 10, "LTC spenddown/waiver oblig. for initial month (eff. date)", section_f_LTC_spenddown_initial_month_checkbox
  EditBox 210, 170, 30, 15, section_f_LTC_spenddown_date
  CheckBox 10, 190, 95, 10, "MA denied (enter eff. date)", section_f_ma_denied_checkbox
  EditBox 210, 185, 30, 15, section_f_ma_denied_date
  CheckBox 10, 205, 180, 10, "MA payment of LTC services denied (enter eff. date)", section_f_ma_payment_denied_checkbox
  EditBox 210, 200, 30, 15, section_f_ma_payment_LTC_date
  CheckBox 10, 220, 10, 10, "", section_f_inelig_for_MA_payment_checkbox
  EditBox 210, 225, 30, 15, section_f_inelig_for_MA_payment_date
  CheckBox 10, 245, 175, 10, "Basic MA continues until specific date (enter date)", section_f_basic_ma_continues_checkbox
  EditBox 210, 240, 30, 15, section_f_basic_ma_continues_date
  CheckBox 10, 260, 185, 15, "Results from asset assess. sent to person (date sent)", section_f_asset_assessment_results_checkbox
  EditBox 210, 255, 30, 15, section_f_results_from_asset_assessment_sent_date
  Text 20, 220, 185, 20, "Person inelig for MA payment of LTSS services until specific date (Enter date inelig. until)"
  ButtonGroup ButtonPressed
    PushButton 220, 290, 50, 15, "Next", next_btn
    PushButton 5, 290, 50, 15, "Previous", previous_btn
  EndDialog
end function
' Dim all functions in variable
Dim section_f_person_applied_MA_checkbox, section_f_dhs_3531_sent_checkbox, section_f_person_applied_date, section_f_dhs_3531_sent_date, section_f_dhs_3543_sent_checkbox, section_f_dhs_3543_sent_date, section_f_dhs_3543_3531_returned_checkbox, section_f_dhs_3543_3531_returned_comments, section_f_dhs_3543_3531_not_returned_checkbox, section_f_ma_opened_checkbox, section_f_ma_opened_date, section_f_basic_ma_medical_spenddown_checkbox, section_f_basic_ma_medical_spenddown, section_f_ma_LTC_services_checkbox, section_f_ma_LTC_services_date, section_f_LTC_spenddown_initial_month_checkbox, section_f_LTC_spenddown_date, section_f_ma_denied_checkbox, section_f_ma_denied_date, section_f_ma_payment_denied_checkbox, section_f_ma_payment_LTC_date, section_f_inelig_for_MA_payment_checkbox, section_f_inelig_for_MA_payment_date, section_f_basic_ma_continues_checkbox, section_f_basic_ma_continues_date, section_f_asset_assessment_results_checkbox, section_f_results_from_asset_assessment_sent_date

'Dialog 9 - Section F: Medical Assistance
function section_f_medical_assistance_changes()
  dialog_count = 9
  section_f_medical_assistance_changes_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "9 - Section F: Medical Assistance - Changes"
    GroupBox 5, 5, 250, 275, "Changes (select all that apply)"
    CheckBox 15, 20, 190, 10, "LTC spenddown/waiver obligation (enter spenddown $)", section_f_LTC_spenddown_checkbox
    EditBox 210, 20, 30, 15, section_f_LTC_spenddown_amount
    CheckBox 15, 40, 10, 10, "", section_f_MA_terminated_checkbox
    Text 25, 40, 170, 20, "MA terminated - basic MA and MA payment of LTSS services (enter eff. date)"
    EditBox 210, 40, 30, 15, section_f_ma_terminated_eff_date
    CheckBox 15, 60, 180, 10, "Basic MA spenddown changed (enter spenddown $)", section_f_basic_ma_spenddown_change_checkbox
    EditBox 210, 60, 30, 15, section_f_basic_ma_spenddown_change_amount
    CheckBox 15, 80, 230, 10, "MA payment of LTSS services terminated; basic MA remains open", section_f_ma_payment_terminated_basic_open_checkbox
    Text 30, 95, 60, 10, "Date terminated:"
    EditBox 90, 90, 30, 15, section_f_ma_payment_terminated_term_date
    Text 140, 95, 70, 10, "Date inelig. through:"
    EditBox 210, 90, 30, 15, section_f_ma_payment_terminated_date_inelig_thru
    CheckBox 15, 110, 145, 10, "Person is deceased (enter date of death)", section_f_person_deceased_checkbox
    EditBox 210, 110, 30, 15, section_f_person_deceased_date_of_death
    CheckBox 15, 125, 110, 10, "Person moved to an institution", section_f_person_moved_institution_checkbox
    EditBox 25, 135, 45, 15, section_f_person_moved_institution_admit_date
    EditBox 70, 135, 95, 15, section_f_person_moved_institution_facility_name
    EditBox 165, 135, 70, 15, section_f_person_moved_institution_phone_number
    EditBox 25, 160, 75, 15, section_f_person_moved_institution_address
    EditBox 100, 160, 75, 15, section_f_person_moved_institution_city
    EditBox 175, 160, 25, 15, section_f_person_moved_institution_state
    EditBox 200, 160, 35, 15, section_f_person_moved_institution_zip
    CheckBox 15, 190, 110, 10, "Person has a new address", section_f_person_new_address_checkbox
    EditBox 25, 200, 75, 15, section_f_person_new_address_date_changed
    EditBox 100, 200, 135, 15, section_f_person_new_address_new_phone_number
    EditBox 25, 225, 75, 15, section_f_person_new_address_address
    EditBox 100, 225, 75, 15, section_f_person_new_address_city
    EditBox 175, 225, 25, 15, section_f_person_new_address_state
    EditBox 200, 225, 35, 15, section_f_person_new_address_zip_code
    CheckBox 15, 250, 135, 10, "Other change (describe reason below)", section_f_other_change_checkbox
    EditBox 25, 260, 225, 15, section_f_person_other_change_description
    ButtonGroup ButtonPressed
      PushButton 220, 290, 50, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 50, 15, "Previous", previous_btn
    Text 30, 150, 205, 10, "Admit date              Facility name                   Phone number"
    Text 30, 175, 205, 10, "Address                                 City                       State       Zip"
    Text 30, 240, 205, 10, "Address                                 City                       State       Zip"
    Text 30, 215, 185, 10, "Date addr. changed        New Phone Number (if changed)"
    ButtonGroup ButtonPressed
      PushButton 220, 290, 50, 15, "Next", next_btn
      CancelButton 270, 290, 50, 15
      PushButton 5, 290, 50, 15, "Previous", previous_btn
    GroupBox 260, 5, 60, 280, "Navigation"
    Text 265, 15, 40, 10, "Section A"
    Text 265, 45, 40, 10, "Section B"
    Text 265, 105, 40, 10, "Section C"
    Text 265, 150, 40, 10, "Section D"
    Text 265, 180, 40, 10, "Section E"
    Text 265, 210, 40, 10, "Section F"
    Text 265, 255, 40, 10, "Section G"
    ButtonGroup ButtonPressed
      PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
' Dim all functions in variable
Dim section_f_LTC_spenddown_checkbox, section_f_LTC_spenddown_amount, section_f_MA_terminated_checkbox,section_f_ma_terminated_eff_date, section_f_basic_ma_spenddown_change_checkbox, section_f_basic_ma_spenddown_change_amount, section_f_ma_payment_terminated_basic_open_checkbox, section_f_ma_payment_terminated_term_date, section_f_ma_payment_terminated_date_inelig_thru, section_f_person_deceased_date_of_death, section_f_person_moved_institution_checkbox, section_f_person_moved_institution_admit_date, section_f_person_moved_institution_facility_name, section_f_person_moved_institution_phone_number, section_f_person_moved_institution_address, section_f_person_moved_institution_city, section_f_person_moved_institution_state, section_f_person_moved_institution_zip, section_f_person_new_address_checkbox, section_f_person_new_address_date_changed, section_f_person_new_address_new_phone_number, section_f_person_new_address_address, section_f_person_new_address_city, section_f_person_new_address_state, section_f_person_new_address_zip_code, section_f_other_change_checkbox, section_f_person_other_change_description, section_f_person_deceased_checkbox

'Dialog 10 - Section G: Comments from eligibility worker
function section_g_comments_elig_worker()
  dialog_count = 10
  section_g_comments_elig_worker_called = True
  BeginDialog Dialog1, 0, 0, 326, 310, "10 - Section G: Comments from elig. worker"
  Text 5, 5, 130, 10, "Enter any additional notes or comments"
  EditBox 5, 15, 225, 15, section_g_elig_comments
  ButtonGroup ButtonPressed
    PushButton 220, 290, 50, 15, "Complete", complete_btn
    CancelButton 270, 290, 50, 15
    PushButton 5, 290, 50, 15, "Previous", previous_btn
  GroupBox 260, 5, 60, 280, "Navigation"
  Text 265, 15, 40, 10, "Section A"
  Text 265, 45, 40, 10, "Section B"
  Text 265, 105, 40, 10, "Section C"
  Text 265, 150, 40, 10, "Section D"
  Text 265, 180, 40, 10, "Section E"
  Text 265, 210, 40, 10, "Section F"
  Text 265, 255, 40, 10, "Section G"
  ButtonGroup ButtonPressed
    PushButton 265, 25, 50, 15, "Contact Info", section_a_contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", section_b_contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
' Dim all variables in function
Dim section_g_elig_comments

Function dialog_selection(dialog_selected) 	'Selects the correct dialog based
  If dialog_selected = 1 then call section_a_contact_info()
  If dialog_selected = 2 then call section_b_assess_results_current_status()
  If dialog_selected = 3 then call section_b_assess_results_initial_assess_case_manager()
  If dialog_selected = 4 then call section_b_assess_results_MA_requests_apps_changes()
  If dialog_selected = 5 then call section_c_comm_elig_worker_exit_reasons()
  If dialog_selected = 6 then call section_c_other_changes_section_d_comments()
  If dialog_selected = 7 then call section_e_contact_info()
  If dialog_selected = 8 then call section_f_medical_assistance()
  If dialog_selected = 9 then call section_f_medical_assistance_changes()
  If dialog_selected = 10 then call section_g_comments_elig_worker()
  If dialog_selected = 11 then call section_a_additional_assessors()
  If dialog_selected = 12 then call section_e_additional_assessors()
End Function

function button_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs.
	If err_msg = "" AND ButtonPressed = next_btn Then dialog_count = dialog_count + 1 'If next is selected, it will go to the next dialog
	If err_msg = "" AND ButtonPressed = previous_btn Then dialog_count = dialog_count - 1	'If previous is selected, it will go to the previous dialog
  If err_msg = "" AND ButtonPressed = -1 then dialog_count = dialog_count + 1   'If enter is pressed, then move to next dialog if no errors
  If err_msg = "" and ButtonPressed = section_a_contact_info_btn then dialog_count = 1
  If err_msg = "" AND ButtonPressed = status_btn then dialog_count = 2
  If err_msg = "" AND ButtonPressed = initial_assessment_btn then dialog_count = 3
  If err_msg = "" AND ButtonPressed = MA_req_app_btn then dialog_count = 4
  If err_msg = "" AND ButtonPressed = exit_reasons_btn then dialog_count = 5
  If err_msg = "" AND ButtonPressed = other_changes_btn then dialog_count = 6
  If err_msg = "" AND ButtonPressed = section_d_comments_btn then dialog_count = 6
  If err_msg = "" AND ButtonPressed = section_b_contact_info_btn then dialog_count = 7
  If err_msg = "" AND ButtonPressed = MA_status_determination_btn then dialog_count = 8
  If err_msg = "" AND ButtonPressed = changes_btn then dialog_count = 9
  If err_msg = "" AND ButtonPressed = section_g_comments_btn then dialog_count = 10
  If err_msg = "" AND ButtonPressed = section_a_add_assessor_btn then dialog_count = 11
  If err_msg = "" AND ButtonPressed = section_e_add_assessor_btn then dialog_count = 12
  If err_msg = "" AND ButtonPressed = section_a_assessor_return_btn then dialog_count = 1
  If ButtonPressed = section_a_assessor_return_no_save_btn then 
    'Reset all Add'l Assessor variables
    section_a_assessor_2 = ""
    section_a_lead_agency_2 = ""
    section_a_phone_number_2 = ""
    section_a_street_address_2 = ""
    section_a_city_2 = ""
    section_a_state_2 = ""
    section_a_zip_code_2 = ""
    section_a_email_address_2 = ""
    section_a_assessor_3 = ""
    section_a_lead_agency_3 = ""
    section_a_phone_number_3 = ""
    section_a_street_address_3 = ""
    section_a_city_3 = ""
    section_a_state_3 = ""
    section_a_zip_code_3 = ""
    section_a_email_address_3 = ""

    dialog_count = 1
  End If 
  If err_msg = "" AND ButtonPressed = section_e_assessor_return_btn then dialog_count = 7
  If ButtonPressed = section_e_assessor_return_no_save_btn then 
    'Reset all Add'l Assessor variables
    section_e_assessor_2 = ""
    section_e_lead_agency_2 = ""
    section_e_phone_number_2 = ""
    section_e_street_address_2 = ""
    section_e_city_2 = ""
    section_e_state_2 = ""
    section_e_zip_code_2 = ""
    section_e_email_address_2 = ""
    section_e_assessor_3 = ""
    section_e_lead_agency_3 = ""
    section_e_phone_number_3 = ""
    section_e_street_address_3 = ""
    section_e_city_3 = ""
    section_e_state_3 = ""
    section_e_zip_code_3 = ""
    section_e_email_address_3 = ""
    
    dialog_count = 7
  End If
  If ButtonPressed = section_a_fill_SWKR_btn Then
    Call navigate_to_MAXIS_screen("STAT", "SWKR")
    'creates a new panel if one doesn't exist, and will needs new if there is not one
    EMReadScreen panel_exists_check, 1, 2, 73
    IF panel_exists_check = "0" THEN
      'If no SWKR panel exists, then msgbox to alert the worker
      msgbox "No SWKR panel exists. Script will return to dialog."
    ELSE
      'Read information from SWKR
      EMReadScreen section_a_assessor, 35, 6, 32
      section_a_assessor = replace(section_a_assessor, "_", "")
      EMReadScreen section_a_street_address, 22, 8, 32
      section_a_street_address = replace(section_a_street_address, "_", "")
      EMReadScreen section_a_city, 15, 10, 32
      section_a_city = replace(section_a_city, "_", "")
      EMReadScreen section_a_phone_number, 14, 12, 34
      section_a_phone_number = replace(section_a_phone_number, " ) ", "")
      section_a_phone_number = replace(section_a_phone_number, " ", "")
      EMReadScreen section_a_state, 2, 10, 54
      EMReadScreen section_a_zip_code, 10, 10, 63
      'Return to STAT/MEMB
      Call navigate_to_MAXIS_screen("STAT", "MEMB")
    END IF
  End If
  If ButtonPressed = section_e_fill_SWKR_btn Then
    Call navigate_to_MAXIS_screen("STAT", "SWKR")
    'creates a new panel if one doesn't exist, and will needs new if there is not one
    EMReadScreen panel_exists_check, 1, 2, 73
    IF panel_exists_check = "0" THEN
      'If no SWKR panel exists, then msgbox to alert the worker
      msgbox "No SWKR panel exists. Script will return to dialog."
    ELSE
      'Read information from SWKR
      EMReadScreen section_e_assessor, 35, 6, 32
      section_e_assessor = replace(section_e_assessor, "_", "")
      EMReadScreen section_e_street_address, 22, 8, 32
      section_e_street_address = replace(section_e_street_address, "_", "")
      EMReadScreen section_e_city, 15, 10, 32
      section_e_city = replace(section_e_city, "_", "")
      EMReadScreen section_e_phone_number, 14, 12, 34
      section_e_phone_number = replace(section_e_phone_number, " ) ", "")
      section_e_phone_number = replace(section_e_phone_number, " ", "")
      EMReadScreen section_e_state, 2, 10, 54
      EMReadScreen section_e_zip_code, 10, 10, 63
      'Return to STAT/MEMB
      Call navigate_to_MAXIS_screen("STAT", "MEMB")
    END IF
    'End at STAT/MEMB
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
  End If

end function

function dialog_specific_error_handling()	'Error handling for main dialog of forms
  'Error handling will display at the point of each dialog and will not let the user continue unless the applicable errors are resolved. Had to list all buttons including -1 so ensure the error reporting is called and hit when the script is run.
  'To do - need these?
	If dialog_count = 11 Then
    If ButtonPressed = -1 Then err_msg = err_msg & vbNewLine & "* You must press either the 'Save Info and Return to Contact Info' or the 'Return to Contact Info WITHOUT Saving Assessor Info' buttons."
  End If

  If dialog_count = 12 Then
    If ButtonPressed = -1 Then err_msg = err_msg & vbNewLine & "* You must press either the 'Save Info and Return to Contact Info' or the 'Return to Contact Info WITHOUT Saving Assessor Info' buttons."
  End If

	If ButtonPressed = section_a_contact_info_btn OR _
    ButtonPressed = status_btn OR _
    ButtonPressed = initial_assessment_btn OR _
    ButtonPressed = MA_req_app_btn OR _
    ButtonPressed = exit_reasons_btn OR _
    ButtonPressed = other_changes_btn OR _
    ButtonPressed = section_d_comments_btn OR _
    ButtonPressed = section_b_contact_info_btn OR _  
    ButtonPressed = MA_status_determination_btn OR _
    ButtonPressed = changes_btn OR _  
    ButtonPressed = section_g_comments_btn OR _
    ButtonPressed = next_btn OR _
    ButtonPressed = previous_btn OR _
    ButtonPressed = instructions_btn OR _
    ButtonPressed = section_a_assessor_return_btn OR _
    ButtonPressed = section_e_assessor_return_btn OR _
    ButtonPressed = -1 Then
      If dialog_count = 1 then 
        If trim(section_a_date_form_sent) = "" OR IsDate(section_a_date_form_sent) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY." 
        If trim(section_a_assessor) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor field." 
        If trim(section_a_lead_agency) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency field." 
        If trim(section_a_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####." 
        If len(trim(section_a_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
        If trim(section_a_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field." 
        If trim(section_a_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field." 
        If trim(section_a_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State field." 
        If len(trim(section_a_state)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State field in the two character format, ex. MN." 
        If trim(section_a_zip_code) = "" or len(trim(section_a_zip_code)) <> 5 Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field in a five number format." 
        If trim(section_a_email_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address field." 
        If hh_memb = "Select One:" Then err_msg = err_msg & vbNewLine & "* You must select the Household Member from the dropdown." 
      End If
      If dialog_count = 11 then 
        If trim(section_a_assessor_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (2) field." 
        If trim(section_a_lead_agency_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (2) field." 
        If trim(section_a_phone_number_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field in the format ###-###-####."
        If len(trim(section_a_phone_number_2)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field in the format ###-###-####."
        If trim(section_a_street_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (2) field." 
        If trim(section_a_city_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (2) field." 
        If trim(section_a_state_2) = "" OR len(trim(section_a_state_2)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State (2) field in the two character format, ex. MN." 
        If trim(section_a_zip_code_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (2) field." 
        If trim(section_a_email_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (2) field." 

        'Handling for Asessor (3) to only trigger errors if some fields are filled in but if completely blank then it will ignore errors
        If trim(section_a_assessor_3) <> "" or trim(section_a_lead_agency_3) <> "" OR trim(section_a_phone_number_3) <> "" OR trim(section_a_street_address_3) <> "" OR trim(section_a_city_3) <> "" OR trim(section_a_state_3) <> "" OR trim(section_a_zip_code_3) <> "" OR trim(section_a_email_address_3) <> "" Then
          If trim(section_a_assessor_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (3) field." 
          If trim(section_a_lead_agency_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (3) field." 
          If trim(section_a_phone_number_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field in the format ###-###-####." 
          If len(trim(section_a_phone_number_3)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field in the format ###-###-####."
          If trim(section_a_street_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (3) field." 
          If trim(section_a_city_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (3) field." 
          If trim(section_a_state_3) = "" OR len(trim(section_a_state_3)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State (3) field in the two character format, ex. MN." 
          If trim(section_a_zip_code_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (3) field." 
          If trim(section_a_email_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (3) field." 
        End If
      End If
      If dialog_count = 12 then 
        If trim(section_e_assessor_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (2) field." 
        If trim(section_e_lead_agency_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (2) field." 
        If trim(section_e_phone_number_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field in the format ###-###-####." 
        If len(trim(section_e_phone_number_2)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field in the format ###-###-####." 
        If trim(section_e_street_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (2) field." 
        If trim(section_e_city_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (2) field." 
        If trim(section_e_state_2) = "" or len(trim(section_e_state_2)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State (2) field in the two character format, ex. MN." 
        If trim(section_e_zip_code_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (2) field." 
        If trim(section_e_email_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (2) field." 

        'Handling for Asessor (3) to only trigger errors if some fields are filled in but if completely blank then it will ignore errors
        If trim(section_e_assessor_3) <> "" or trim(section_e_lead_agency_3) <> "" OR trim(section_e_phone_number_3) <> "" OR trim(section_e_street_address_3) <> "" OR trim(section_e_city_3) <> "" OR trim(section_e_state_3) <> "" OR trim(section_e_zip_code_3) <> "" OR trim(section_e_email_address_3) <> "" Then
          If trim(section_e_assessor_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (3) field." 
          If trim(section_e_lead_agency_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (3) field." 
          If trim(section_e_phone_number_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field in the format ###-###-####." 
          If len(trim(section_e_phone_number_3)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field in the format ###-###-####." 
          If trim(section_e_street_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (3) field." 
          If trim(section_e_city_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (3) field." 
          If trim(section_e_state_3) = "" or len(trim(section_e_state_3)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State (3) field in the two character format, ex. MN." 
          If trim(section_e_zip_code_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (3) field." 
          If trim(section_e_email_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (3) field." 
        End If
      End If
      If dialog_count = 2 then 
        If section_g_person_requesting_already_enrolled_LTC + section_g_person_will_reside_institution_checkbox = 0 Then err_msg = err_msg & vbNewLine & "* You must check one of the boxes for the person's current status."
        If section_g_person_requesting_already_enrolled_LTC + section_g_person_will_reside_institution_checkbox = 2 Then err_msg = err_msg & vbNewLine & "* Only select the second option for the person's current status if both options apply."

        If section_b_program_type = "Select one:" Then err_msg = err_msg & vbNewLine & "* You must select the program the person is requesting or is currently enrolled in from the dropdown list." 
        If section_b_diversion_checkbox + section_b_conversion_checkbox = 2 Then err_msg = err_msg & vbNewLine & "* You can only select one checkbox option - Diversion or Conversion."
        If trim(section_b_admission_date) = "" or IsDate(section_b_admission_date) = False Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY."
        If trim(section_b_facility) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Facility field."
        If trim(section_b_institution_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
        If len(trim(section_b_institution_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
        If trim(section_b_institution_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field."
        If trim(section_b_institution_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field."
        If trim(section_b_institution_state) = "" OR len(trim(section_b_institution_state)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State field in the two character format, ex. MN."
        If trim(section_b_institution_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field."
      End if 
      If dialog_count = 3 then 
        If trim(section_b_assessment_date) = "" or IsDate(section_b_assessment_date) = False Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessment Date field in the format MM/DD/YYYY."
        If section_b_assessment_determination = "Select one:" Then err_msg = err_msg & vbNewLine & "* You must select the Assessment Determination from the dropdown list." 
        If section_b_open_to_waiver_yes_checkbox + section_b_open_to_waiver_no_checkbox = 2 Then err_msg = err_msg & vbNewLine & "* You can only select one checkbox option for whether the person will open to waiver/AC/ECS - Yes or No."
        If section_b_open_to_waiver_yes_checkbox + section_b_open_to_waiver_no_checkbox = 0 Then err_msg = err_msg & vbNewLine & "* You must select one checkbox option for whether the person will open to waiver/AC/ECS - Yes or No."
        If trim(section_b_monthly_waiver_costs) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the estimated monthly waiver/AC costs field."
        If trim(section_b_waiver_effective_date) = "" or IsDate(section_b_waiver_effective_date) = False Then err_msg = err_msg & vbNewLine & "* You must fill out the anticipated effective date field in the format MM/DD/YYYY."
        If section_b_yes_case_manager + section_b_yes_someone_else_case_manager + section_b_no_case_manager > 1 Then err_msg = err_msg & vbNewLine & "* You can only select one checkbox for whether the person has a case manager."
        If section_b_no_case_manager = 1 Then 
          If trim(section_b_case_manager_name) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Case Manager Name field."
          If trim(section_b_case_manager_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field for the case manager in the format ###-###-####."
          If len(trim(section_b_case_manager_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field for the case manager in the format ###-###-####."
        End If
      End if 
      If dialog_count = 4 then 
        'To do - Handling needed?
        ' section_b_applied_MA_LTC_checkbox
        If section_b_ma_enrollee_checkbox = 1 Then
          If trim(section_b_date_dhs_3543_provided) = "" or IsDate(section_b_date_dhs_3543_provided) = False Then err_msg = err_msg & vbNewLine & "* You must enter the date the assessor provided the DHS-3543."
        End If
        'To do - handling needed?
        If section_b_completed_dhs_3543_3531_attached_checkbox = 1 Then
          If trim(section_b_dhs_3543_3531_sent_to_county_date) = "" or IsDate(section_b_dhs_3543_3531_sent_to_county_date) = False Then err_msg = err_msg & vbNewLine & "* You must enter the date the assessor provided the DHS-3543."
        End If
        'Tod do - handling needed?
        ' section_b_send_dhs_3543_checkbox
        If section_b_send_dhs_3531_checkbox = 1 Then
          If trim(section_b_dhs_3543_3531_sent_to_county_date) = "" or IsDate(section_b_dhs_3543_3531_sent_to_county_date) = False Then err_msg = err_msg & vbNewLine & "* You must enter the date the assessor provided the DHS-3543."
        End If
        If section_b_send_dhs_3340_checkbox = 1 Then
          If trim(section_b_send_dhs_3340_address) = "" or trim(section_b_send_dhs_3340_city) = "" or trim(section_b_send_dhs_3340_state) = "" or trim(section_b_send_dhs_3340_zip) = "" Then err_msg = err_msg & vbNewLine & "* The checkbox for Send DHS-3340 to person (asset assessment needed) is checked so you must fill out the Address, City, State, and Zip Code fields below the checkbox."
        End If          
        If section_b_person_no_longer_institutional_LOC_checkbox = 1 Then
          If trim(section_b_date_waiver_exit) = "" OR IsDate(section_b_date_waiver_exit) = False Then err_msg = err_msg & vbNewLine & "* The checkbox for Person no longer meets institutional LOC is checked. You must enter the effective date of waiver exit."
        End If 
        If section_b_person_enroll_another_program = 1 Then
          If section_b_enroll_another_program_list = "Select one:" Then err_msg = err_msg & vbNewLine & "* The checkbox for Person chooses to enroll in another program. You must select the program from the dropdown list."
        End If 
      End if 
      If dialog_count = 5 then 
        If section_c_exited_waiver_program_checkbox = 1 Then
          If trim(section_c_date_waiver_exit) = "" or IsDate(section_c_date_waiver_exit) = False Then err_msg = err_msg & vbNewLine & "* You must enter the effective date of the waiver exit."
        End If
        'To do - handling needed?
        ' section_c_hospital_admission_checkbox, section_c_nursing_facility_admission_checkbox, section_c_person_informed_choice_checkbox, section_c_residential_treatment_admission_checkbox
        If section_c_person_deceased_checkbox = 1 Then
          If trim(section_c_date_of_death) = "" or IsDate(section_c_date_of_death) = False Then err_msg = err_msg & vbNewLine & "* You must fill out the date of death field in the format MM/DD/YYYY."
        End If
        If section_c_person_moved_out_of_state_checkbox = 1 Then
          If trim(section_c_date_of_move) = "" or IsDate(section_c_date_of_move) = False Then err_msg = err_msg & vbNewLine & "* You must fill out the date of move field in the format MM/DD/YYYY."
        End If
        If section_c_exited_for_other_reasons_checkbox = 1 Then
          If trim(section_c_exited_for_other_reasons_explanation) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Exited for other reasons field."
        End If
      End if 
      If dialog_count = 6 then 
        If section_c_program_type_list = "Select one:" Then err_msg = err_msg & vbNewLine & "* You must select the program type from the dropdown list."

        If section_c_diversion_checkbox + section_c_conversion_checkbox = 2 Then err_msg = err_msg & vbNewLine & "* You can only select one option, not both, for Diversion or Conversion."
        
        If section_c_person_moved_new_address_checkbox = 1 Then
          If trim(section_c_date_address_changed) = "" or IsDate(section_c_date_address_changed) = False Then err_msg = err_msg & vbNewLine & "* You must enter the Date of Address Change in the format MM/DD/YYYY."
          If trim(section_c_street_address) = "" OR trim(section_c_city) = "" or trim(section_c_state) = "" OR trim(section_c_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the fields for the new address (address, state, city, and zip code)."
        End If
        If section_c_new_legal_rep_checkbox = 1 Then
          If trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the First Name, Last Name, and Phone Number fields for the new legal representative."
          If len(trim(section_c_legal_rep_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
          If trim(section_c_legal_rep_street_address) = "" or trim(section_c_legal_rep_city) = "" or trim(section_c_legal_rep_state) = "" OR trim(section_c_legal_rep_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address, City, State, and Zip Code fields for the new legal representative."
        End If
        If section_c_person_return_to_community_checkbox = 1 Then 
          If trim(section_c_qual_admission_eff_date) = "" OR IsDate(section_c_qual_admission_eff_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Effective Date for the Person returning to community w/in 121 days of a qual. admission."
        End If
        If section_c_other_changes_program_checkbox = 1 Then
          If trim(section_c_other_changes_program) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the field to describe the Other changes related to program/service eligibility."
        End If 
        'To do - handling needed?
        ' section_d_additional_comments
      End if 
      If dialog_count = 7 then 
        If trim(section_e_date_form_sent) = "" OR IsDate(section_e_date_form_sent) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY." 
        If trim(section_e_assessor) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor field." 
        If trim(section_e_lead_agency) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency field." 
        If trim(section_e_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
        If len(trim(section_e_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field in the format ###-###-####."
        If trim(section_e_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field." 
        If trim(section_e_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field." 
        If trim(section_e_state) = "" or len(trim(section_e_state)) <> 2 Then err_msg = err_msg & vbNewLine & "* You must fill out the State field in the two character format, ex. MN." 
        If trim(section_e_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field." 
        If trim(section_e_email_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address field." 
        If hh_memb = "Select One:" Then err_msg = err_msg & vbNewLine & "* You must select the Household Member from the dropdown." 
      End if 
      If dialog_count = 8 then 
        If section_f_person_applied_MA_checkbox = 1 Then
          If trim(section_f_person_applied_date) = "" OR IsDate(section_f_person_applied_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the person applied for MA/MA-LTC in the format MM/DD/YYYY."
        End If
        If section_f_dhs_3531_sent_checkbox = 1 Then
          If trim(section_f_dhs_3531_sent_date) = "" OR IsDate(section_f_dhs_3531_sent_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the DHS-3531 was sent to the person in the format MM/DD/YYYY."
        End If
        If section_f_dhs_3543_sent_checkbox = 1 Then
          If trim(section_f_dhs_3543_sent_date) = "" OR IsDate(section_f_dhs_3543_sent_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the DHS-3543 was sent to the person in the format MM/DD/YYYY."
        End If
        'To do - handling needed?
        ' section_f_dhs_3543_3531_returned_checkbox, section_f_dhs_3543_3531_returned_comments
        ' section_f_dhs_3543_3531_not_returned_checkbox 
        If section_f_ma_opened_checkbox = 1 Then
          If trim(section_f_ma_opened_date) = "" OR IsDate(section_f_ma_opened_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the DHS-3543 was sent to the person in the format MM/DD/YYYY."
        End If
        If section_f_basic_ma_medical_spenddown_checkbox = 1 Then
          If trim(section_f_basic_ma_medical_spenddown) = "" Then err_msg = err_msg & vbNewLine & "* You must enter the dollar amount in the basic MA medical spenddown field."
        End If
        If section_f_ma_LTC_services_checkbox = 1 Then
          If trim(section_f_ma_LTC_services_date) = "" OR IsDate(section_f_ma_LTC_services_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the effective date for when the MA for LTC services opened in the format MM/DD/YYYY."
        End If
        If section_f_LTC_spenddown_initial_month_checkbox = 1 Then
          If trim(section_f_LTC_spenddown_date) = "" OR IsDate(section_f_LTC_spenddown_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the effective date for the LTC spenddown/waiver obligation for initial month in the format MM/DD/YYYY."
        End If
        If section_f_ma_denied_checkbox = 1 Then
          If trim(section_f_ma_denied_date) = "" OR IsDate(section_f_ma_denied_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the effective date for the MA denial in the format MM/DD/YYYY."
        End If
        If section_f_ma_payment_denied_checkbox = 1 Then
          If trim(section_f_ma_payment_LTC_date) = "" OR IsDate(section_f_ma_payment_LTC_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the effective date for the MA payment of LTC services denial in the format MM/DD/YYYY."
        End If
        If section_f_inelig_for_MA_payment_checkbox = 1 Then
          If trim(section_f_inelig_for_MA_payment_date) = "" OR IsDate(section_f_inelig_for_MA_payment_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You checked the box that the person is ineligible for MA payment of LTSS services until a specific date. You must enter the date the ineligibility lasts until in the format MM/DD/YYYY."
        End If
        If section_f_basic_ma_continues_checkbox = 1 Then
          If trim(section_f_basic_ma_continues_date) = "" OR IsDate(section_f_basic_ma_continues_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date that basic MA continues until in the format MM/DD/YYYY."
        End If
        If section_f_asset_assessment_results_checkbox = 1 Then
          If trim(section_f_results_from_asset_assessment_sent_date) = "" OR IsDate(section_f_results_from_asset_assessment_sent_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the results from the asset assessment were sent to the person in the format MM/DD/YYYY."
        End If
      End If
      If dialog_count = 9 then
        If section_f_LTC_spenddown_checkbox = 1 Then
          If trim(section_f_LTC_spenddown_amount) = "" Then err_msg = err_msg & vbNewLine & "* You must enter the spenddown dollar amount for the LTC spenddown/waiver obligation."
        End If
        If section_f_MA_terminated_checkbox = 1 Then
          If trim(section_f_ma_terminated_eff_date) = "" OR IsDate(section_f_ma_terminated_eff_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the effective date for the MA termination for basic MA and MA payment of LTSS services in the format MM/DD/YYYY."
        End If
        If section_f_basic_ma_spenddown_change_checkbox = 1 Then
          If trim(section_f_basic_ma_spenddown_change_amount) = "" Then err_msg = err_msg & vbNewLine & "* You must enter the spenddown dollar amount for the basic MA spenddown."
        End If
        If section_f_ma_payment_terminated_basic_open_checkbox = 1 Then
          If trim(section_f_ma_payment_terminated_term_date) = "" OR IsDate(section_f_ma_payment_terminated_term_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the termination date of the MA payment of LTSS services in the format MM/DD/YYYY."
          If trim(section_f_ma_payment_terminated_date_inelig_thru) = "" OR IsDate(section_f_ma_payment_terminated_date_inelig_thru) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date the ineligibility lasts through in the format MM/DD/YYYY."
        End If
        If section_f_person_deceased_checkbox = 1 Then
          If trim(section_f_person_deceased_date_of_death) = "" OR IsDate(section_f_person_deceased_date_of_death) = FALSE Then err_msg = err_msg & vbNewLine & "* You must enter the date of death in the format MM/DD/YYYY."
        End If
        If section_f_person_moved_institution_checkbox = 1 Then
          If trim(section_f_person_moved_institution_admit_date) = "" OR IsDate(section_f_person_moved_institution_admit_date) = FALSE Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date in the format MM/DD/YYYY."
          If trim(section_f_person_moved_institution_facility_name) = "" OR trim(section_f_person_moved_institution_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date, facility name, and phone number for the institution."
          If len(trim(section_f_person_moved_institution_phone_number)) <> 12 Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the phone number for the institution in the format ###-###-####."
          If trim(section_f_person_moved_institution_address) = "" OR trim(section_f_person_moved_institution_city) = "" OR trim(section_f_person_moved_institution_state) = "" OR trim(section_f_person_moved_institution_zip) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the address, city, state, and zip code for the institution."
        End If
        If section_f_person_new_address_checkbox = 1 Then
          If trim(section_f_person_new_address_date_changed) = "" OR IsDate(section_f_person_new_address_date_changed) = False Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person has a new address. You must enter the date of the address change in the format MM/DD/YYYY."
          If trim(section_f_person_new_address_address) = "" OR trim(section_f_person_new_address_city) = "" OR trim(section_f_person_new_address_state) = "" OR trim(section_f_person_new_address_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person has a new address. You must enter the new address, city, state, and zip code for the new address."
        End If
        If section_f_other_change_checkbox = 1 Then
          If trim(section_f_person_other_change_description) = "" Then err_msg = err_msg & vbNewLine & "* You checked the Other change box. You must describe the change in the field provided."
        End If
      End If
      ' If dialog_count = 10 then
      '   'No error handling needed for comments
      ' End If
  End If
  If ButtonPressed = complete_btn Then
    If section_a_contact_info_called = False OR _
    section_b_assess_results_current_status_called = False OR _
    section_b_assess_results_MA_requests_apps_changes_called = False OR _
    section_c_comm_elig_worker_exit_reasons_called = False OR _
    section_c_other_changes_section_d_comments_called = False OR _
    section_e_contact_info_called = False OR _
    section_f_medical_assistance_called = False OR _
    section_f_medical_assistance_changes_called = False OR _
    section_g_comments_elig_worker_called = False Then
      err_msg = err_msg & vbNewLine & "* All dialogs must be viewed/completed. Please review the following dialogs:"
    End If

    If section_a_contact_info_called = False Then err_msg = err_msg & vbNewLine & "--> Section A: Contact Info"
    If section_b_assess_results_current_status_called = False Then err_msg = err_msg & vbNewLine & "--> Section B: Status"
    If section_b_assess_results_MA_requests_apps_changes_called = False Then err_msg = err_msg & vbNewLine & "--> Section B: Initial Assess."
    If section_c_comm_elig_worker_exit_reasons_called = False Then err_msg = err_msg & vbNewLine & "--> Section C: Exit Reasons"
    If section_c_other_changes_section_d_comments_called = False Then err_msg = err_msg & vbNewLine & "--> Section C: Other Changes"
    If section_e_contact_info_called = False Then err_msg = err_msg & vbNewLine & "--> Section E: Contact Info"
    If section_f_medical_assistance_called = False Then err_msg = err_msg & vbNewLine & "--> Section F: MA Status/Det"
    If section_f_medical_assistance_changes_called = False Then err_msg = err_msg & vbNewLine & "--> Section F: Changes"
    If section_g_comments_elig_worker_called = False Then err_msg = err_msg & vbNewLine & "--> Section G: Comments"
  End If
	If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
end function

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number and footer month/year
EMConnect ""
call check_for_MAXIS(False) 'Checking to see that we're in MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)

'Initial Dialog - Instructions
Dialog1 = "" 'Blanking out previous dialog detail
'Showing the case number - defining the dialog for the case number
BeginDialog Dialog1, 0, 0, 221, 95, "Enter LTC-5181 Form Details"
  Text 10, 5, 200, 20, "Script Purpose: Enter details from submitted LTC-5181 form. Creates a CASE/NOTE with form details."
  Text 20, 30, 50, 10, "Case Number:"
  EditBox 75, 25, 50, 15, MAXIS_case_number
  Text 10, 50, 60, 10, "Worker Signature:"
  EditBox 75, 45, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 75, 45, 15
    CancelButton 170, 75, 45, 15
    PushButton 150, 25, 65, 15, "Script Instructions", instructions_btn
EndDialog

DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog Dialog1				'main dialog
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
If is_this_priv = True then script_end_procedure("Case is privileged. The script will now end.")
'Create list of HH members
Call Generate_Client_List(HH_Memb_DropDown, "Select One:")

'Start at the first dialog
dialog_count = 1

Do
	Do
		Do
			Dialog1 = "" 'Blanking out previous dialog detail
      Call dialog_selection(dialog_count)

      'Blank out variables on each new dialog
			err_msg = ""

			dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
			cancel_confirmation
      Call dialog_specific_error_handling	'function for error handling of main dialog of forms
			Call button_movement()				'function to move throughout the dialogs
		Loop until err_msg = ""
	Loop until ButtonPressed = complete_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call check_for_MAXIS(False) 'Checking to see that we're in MAXIS
'Start at SELF
Call back_to_SELF

'ACTIONS----------------------------------------------------------------------------------------------------
'Read through panels to determine status and updates if needed
'Update ADDR
'--Fields with new address
new_address_provided = False
section_c_section_f_both_new_addresses = False
section_c_person_moved_new_address_only = False
section_f_person_new_address_only = False
date_of_death_provided = False

'--Fields with new assessor details (SWKR panel)
'Navigate to STAT/SWKR to gather details
Call navigate_to_MAXIS_screen("STAT", "SWKR")
EmReadScreen swkr_does_not_exist, 19, 24, 2
If swkr_does_not_exist = "SWKR DOES NOT EXIST" Then
  swkr_panel_exists = False
Else
  swkr_panel_exists = True
  'Read the SWKR screen - name, street, city, state, zip, phone
  EmReadScreen current_swkr_name, 35, 6, 32
  current_swkr_name = replace(current_swkr_name, "_", "")
  EmReadScreen current_swkr_street, 22, 8, 32
  current_swkr_street = replace(current_swkr_street, "_", "")
  EmReadScreen current_swkr_city, 15, 10, 32
  current_swkr_city = replace(current_swkr_city, "_", "")
  EmReadScreen current_swkr_state, 2, 10, 54
  EmReadScreen current_swkr_zip, 5, 10, 63
  EmReadScreen current_swkr_area_code, 3, 12, 34
  EmReadScreen current_swkr_prefix_code, 3, 12, 40
  EmReadScreen current_swkr_line_code, 4, 12, 44
  current_swkr_phone_number = current_swkr_area_code & current_swkr_prefix_code & current_swkr_line_code
  current_swkr_panel_info = current_swkr_name & "(" & current_swkr_street & ", " & current_swkr_city & ", " & current_swkr_state & " " & current_swkr_zip & "; " & current_swkr_phone_number & ")"
End If

If section_c_person_moved_new_address_checkbox = 1 OR section_f_person_new_address_checkbox = 1 Then
    new_address_provided = True
    'Navigate to STAT/ADDR
    Call navigate_to_MAXIS_screen("STAT", "ADDR")
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

    addr_eff_date = replace(addr_eff_date, " ", "/")

    current_ADDR_address = addr_eff_date & "; " & resi_street_full & ", " & resi_state & ", " & resi_zip & " (" & "County: " & resi_county & "; " & "Ver: " & addr_verif & "; " & "Living Sit: " & addr_living_sit & ")"

    'If both addresses have been added, then need to compare them to determine if they match
    If section_c_person_moved_new_address_checkbox = 1 AND section_f_person_new_address_checkbox = 1 Then
      section_c_section_f_both_new_addresses = True
      section_c_person_moved_new_address_full = UCase(section_c_street_address & ", " & section_c_city & ", " & section_c_state & " " & section_c_zip_code)
      section_f_person_new_address_full = UCase(section_f_person_new_address_address & ", " & section_f_person_new_address_city & ", " & section_f_person_new_address_state & " " & section_f_person_new_address_zip_code)
      If section_c_person_moved_new_address_full = section_f_person_new_address_full Then
        section_c_section_f_addresses_match = True
      Else
        section_c_section_f_addresses_match = False
      End If
    ElseIf section_c_person_moved_new_address_checkbox = 1 AND section_f_person_new_address_checkbox <> 1 Then
      'Only section c new address checked
      section_c_person_moved_new_address_only = True
      section_c_person_moved_new_address_full = UCase(section_c_street_address & ", " & section_c_city & ", " & section_c_state & " " & section_c_zip_code)
    ElseIf section_c_person_moved_new_address_checkbox <> 1 AND section_f_person_new_address_checkbox = 1 Then
      'Only section f new address checked
      section_f_person_new_address_only = True
      section_f_person_new_address_full = UCase(section_f_person_new_address_address & ", " & section_f_person_new_address_city & ", " & section_f_person_new_address_state & " " & section_f_person_new_address_zip_code)
    End If
End If

msgbox "section_c_section_f_addresses_match > " & section_c_section_f_addresses_match & vbcr & "section_c_person_moved_new_address_only > " & section_c_person_moved_new_address_only & vbcr & "section_f_person_new_address_only > " & section_f_person_new_address_only 

' 'Display ADDR information from panel and entered information
' ' section_c_person_moved_new_address_checkbox, section_c_date_address_changed, section_c_street_address, section_c_city, section_c_state, section_c_zip_code

' ' section_f_person_new_address_checkbox, section_f_person_new_address_date_changed, section_f_person_new_address_new_phone_number, section_f_person_new_address_address, section_f_person_new_address_city, section_f_person_new_address_state, section_f_person_new_address_zip_code

' 'Dialog that will display details from current SWKR, as well as assessors added in the dialogs
' ' section_a_assessor, section_a_lead_agency, section_a_phone_number, section_a_street_address, section_a_city, section_a_state, section_a_zip_code, section_a_email_address, hh_memb

' ' section_a_assessor_2, section_a_lead_agency_2, section_a_phone_number_2, section_a_street_address_2, section_a_city_2, section_a_state_2, section_a_zip_code_2, section_a_email_address_2, section_a_assessor_3, section_a_lead_agency_3, section_a_phone_number_3, section_a_street_address_3, section_a_city_3, section_a_state_3, section_a_zip_code_3, section_a_email_address_3

' ' section_e_date_form_sent, section_e_assessor, section_e_lead_agency, section_e_phone_number, section_e_street_address, section_e_city, section_e_state, section_e_zip_code, section_e_email_address, hh_memb

' ' section_e_assessor_2, section_e_lead_agency_2, section_e_phone_number_2, section_e_street_address_2, section_e_city_2, section_e_state_2, section_e_zip_code_2, section_e_email_address_2, section_e_assessor_3, section_e_lead_agency_3, section_e_phone_number_3, section_e_street_address_3, section_e_city_3, section_e_state_3, section_e_zip_code_3, section_e_email_address_3

'--Fields with date of death
'Navigate to STAT/MEMB to gather details
If section_c_person_deceased_checkbox = 1 OR section_f_person_deceased_checkbox = 1 Then
  date_of_death_provided = True
  Call navigate_to_MAXIS_screen("STAT", "MEMB")
  'Navigate to HH Memb
  Call write_value_and_transmit(left(hh_memb, 2), 20, 76)
  EMReadScreen memb_date_of_death, 10, 19, 42
  If memb_date_of_death = "__ __ ____" then 
    memb_panel_date_of_death_exists = False
  Else
    memb_panel_date_of_death_exists = True
    memb_date_of_death = replace(memb_date_of_death, " ", "/")
  End If
  'If both addresses have been added, then need to compare them to determine if they match
  If section_c_person_deceased_checkbox = 1 AND section_f_person_deceased_checkbox = 1 Then
    section_c_section_f_both_new_DOD = True
    'Convert both dates of death to dates to compare them
    section_c_date_of_death = DateAdd("m", 0, section_c_date_of_death)
    section_f_person_deceased_date_of_death = dateadd("m", 0, section_f_person_deceased_date_of_death)
    If section_c_date_of_death = section_f_person_deceased_date_of_death Then
      'The dates are the same date
      section_c_section_f_dates_of_death_match = True
    Else
      'The dates do not match - this shouldn't happen
      section_c_section_f_dates_of_death_match = False
    End If
  ElseIf section_c_person_deceased_checkbox = 1 AND section_f_person_deceased_checkbox <> 1 Then
    'If only section c DOD entered
    section_c_person_deceased_only = True
  ElseIf section_c_person_deceased_checkbox <> 1 AND section_f_person_deceased_checkbox = 1 Then
    'If only section f DOD entered
    section_f_person_deceased_only = True
  End If
End If

' section_c_person_deceased_checkbox, section_c_date_of_death
' section_f_person_deceased_date_of_death, section_f_person_deceased_checkbox

'Format
'x, y, length, height
'Set starting position for y_pos
y_pos = 60

'Calculate size of the SWKR group box
swkr_group_box = 95
If section_a_assessor_2 <> "" Then swkr_group_box = swkr_group_box + 10
If section_a_assessor_3 <> "" Then swkr_group_box = swkr_group_box + 10
If section_e_assessor_2 <> "" Then swkr_group_box = swkr_group_box + 10 
If section_e_assessor_3 <> "" Then swkr_group_box = swkr_group_box + 10 

'ADDR, SWKR, DOD match
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 325, 310, "STAT Panel Updates"
  GroupBox 5, 5, 315, swkr_group_box, "Update SWKR"
  Text 15, 20, 70, 10, "Current SWKR Panel: "
  If swkr_panel_exists = True Then Text 15, 30, 290, 10, current_swkr_panel_info
  If swkr_panel_exists = False Then Text 15, 30, 290, 10, "No SWKR Panel Exists" 
  CheckBox 15, 45, 290, 10, "Check here to update the SWKR panel (select ONE Assessor to use for update below):", swkr_update_checkbox
  CheckBox 25, 60, 275, 10, "Section A - Assessor 1: " & section_a_assessor, section_a_assessor_1_checkbox
  If section_a_assessor_2 <> "" Then 
    CheckBox 25, y_pos + 10, 275, 10, "Section A - Assessor 2: " & section_a_assessor_2, section_a_assessor_2_checkbox
    y_pos = y_pos + 10
  End If
  If section_a_assessor_3 <> "" Then 
    CheckBox 25, y_pos + 10, 275, 10, "Section A - Assessor 3: " & section_a_assessor_3, section_a_assessor_3_checkbox
    y_pos = y_pos + 10
  End If
  CheckBox 25, y_pos + 10, 275, 10, "Section E - Assessor 1: " & section_e_assessor, section_e_assessor_1_checkbox
  y_pos = y_pos + 10
  If section_e_assessor_2 <> "" Then 
    CheckBox 25, y_pos + 10, 275, 10, "Section E - Assessor 2: " & section_e_assessor_2, section_e_assessor_2_checkbox
    y_pos = y_pos + 10
  End If
  If section_e_assessor_3 <> "" Then 
    CheckBox 25, y_pos + 10, 275, 10, "Section E - Assessor 3: " & section_e_assessor_3, section_e_assessor_3_checkbox
    y_pos = y_pos + 10
  End If
  Text 10, y_pos + 15, 145, 10, "All notices to Social Worker (select Y or N):"
  CheckBox 160, y_pos + 15, 25, 10, "Yes", notices_to_social_worker_y_checkbox
  CheckBox 190, y_pos + 15, 25, 10, "No", notices_to_social_worker_n_checkbox
  y_pos = y_pos + 30
  'Add the addresses if needed
  If new_address_provided = True Then
    'Insert current ADDR information
    GroupBox 5, y_pos, 315, 75, "Update ADDR"
    Text 15, y_pos + 10, 70, 10, "Current ADDR Panel: "
    Text 15, y_pos + 20, 290, 20, current_ADDR_address
    y_pos = y_pos + 30
    If section_c_section_f_both_new_addresses = True Then
      If section_c_section_f_addresses_match = True Then
        'The addresses DO match so only need to display one of them
        CheckBox 15, y_pos + 10, 285, 10, "Check here to update the ADDR panel", addr_update_checkbox_section_c_section_f_match
        Text 25, y_pos + 25, 290, 10, "New Address Entered on LTC-5181: " & section_f_person_new_address_full
        y_pos = y_pos + 35
      ElseIf section_c_section_f_addresses_match = False Then
        'The addresses DO NOT match so need to display both of them
        CheckBox 15, y_pos + 10, 285, 10, "Check here to update the ADDR panel (select ONE address to use for update below):", addr_update_multiple_checkbox
        CheckBox 25, y_pos + 20, 275, 10, "Section C - New Address: " & section_c_person_moved_new_address_full, multiple_section_c_new_address_checkbox
        CheckBox 25, y_pos + 30, 275, 10, "Section F - New Address: " & section_f_person_new_address_full, multiple_section_f_new_address_checkbox
        y_pos = y_pos + 30
      End If
    ElseIf section_c_person_moved_new_address_only = True Then
      'Only need to display the section c address
      CheckBox 15, y_pos + 10, 285, 10, "Check here to update the ADDR panel", addr_update_checkbox_section_c
      Text 25, y_pos + 25, 290, 10, "New Address Entered on LTC-5181: " & section_c_person_moved_new_address_full
      y_pos = y_pos + 35
    ElseIf section_f_person_moved_new_address_only = True Then
      'Only need to display the section f address
      CheckBox 15, y_pos + 10, 285, 10, "Check here to update the ADDR panel", addr_update_checkbox_section_f
      Text 25, y_pos + 25, 290, 10, "New Address Entered on LTC-5181: " & section_f_person_new_address_full
      y_pos = y_pos + 35
    End If
  End If
  If date_of_death_provided = True then
    y_pos = y_pos + 20
    GroupBox 5, y_pos, 315, 65, "Update MEMB (Date of Death)"
    Text 15, y_pos + 15, 100, 10, "Current DOD on MEMB Panel: "
    If memb_panel_date_of_death_exists = True Then Text 120, y_pos + 15, 100, 10, memb_date_of_death
    If memb_panel_date_of_death_exists = False Then Text 120, y_pos + 15, 100, 10, "No date of death entered"
    y_pos = y_pos + 15 
    If section_c_section_f_both_new_DOD = True Then
      If section_c_section_f_dates_of_death_match = True Then
        CheckBox 15, y_pos + 15, 275, 10, "Check here to update the date of death on MEMB panel (select ONE DOD below):", date_of_death_update_multiple_checkbox
        CheckBox 25, y_pos + 25, 275, 10, "Section C - Date of Death: " & section_c_date_of_death, section_c_date_of_death_checkbox
        CheckBox 25, y_pos + 35, 275, 10, "Section F - Date of Death: " & section_f_person_deceased_date_of_death, section_f_date_of_death_checkbox
        y_pos = y_pos + 10
      ElseIf section_c_section_f_dates_of_death_match = False Then
        CheckBox 15, y_pos + 15, 275, 10, "Check here to update the date of death on MEMB panel", date_of_death_update_checkbox
        Text 25, y_pos + 30, 290, 10, "Date of Death Entered on LTC-5181: " & section_c_date_of_death
        y_pos = y_pos + 10
      End If
    ElseIf section_c_person_deceased_only = True Then
      CheckBox 15, y_pos + 15, 275, 10, "Check here to update the date of death on MEMB panel", date_of_death_update_checkbox
      Text 25, y_pos + 30, 290, 10, "Date of Death Entered on LTC-5181: " & section_c_date_of_death
      y_pos = y_pos + 10
    ElseIf section_f_person_deceased_only = True Then
      CheckBox 15, y_pos + 15, 275, 10, "Check here to update the date of death on MEMB panel", date_of_death_update_checkbox
      Text 25, y_pos + 30, 290, 10, "Date of Death Entered on LTC-5181: " & section_f_person_deceased_date_of_death
      y_pos = y_pos + 10
    End If
  End If
  Text 5, 295, 135, 10, "Enter footer month and year for updates:"
  EditBox 145, 290, 25, 15, footer_month_updates
  EditBox 175, 290, 25, 15, footer_year_updates
  ButtonGroup ButtonPressed
    PushButton 205, 290, 60, 15, "UPDATE Panels", update_panels_btn
    PushButton 265, 290, 55, 15, "SKIP Updates", skip_panel_updates_btn
EndDialog

'Dialog validation
Do
  Do
    'Blank out variables on each new dialog
    err_msg = ""

    dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
    cancel_confirmation

    'Error handling only if worker intends to update panels
    If ButtonPressed = update_panels_btn Then
      If swkr_update_checkbox = 1 AND ((section_a_assessor_1_checkbox + section_a_assessor_2_checkbox + section_a_assessor_3_checkbox + section_e_assessor_1_checkbox + section_e_assessor_2_checkbox + section_e_assessor_3_checkbox = 0) OR (section_a_assessor_1_checkbox + section_a_assessor_2_checkbox + section_a_assessor_3_checkbox + section_e_assessor_1_checkbox + section_e_assessor_2_checkbox + section_e_assessor_3_checkbox > 1)) Then err_msg = err_msg & vbCr & "* If you want to update the SWKR panel, you must select ONLY one assessor to use for the update." 

      If addr_update_multiple_checkbox = 1 AND ((multiple_section_c_new_address_checkbox + multiple_section_f_new_address_checkbox = 0) OR (multiple_section_c_new_address_checkbox + multiple_section_f_new_address_checkbox = 2)) Then err_msg = err_msg & vbCr & "* If you want to update the ADDR panel, you must select ONLY one address to use for the update." 

      If date_of_death_update_multiple_checkbox = 1 AND ((section_c_date_of_death_checkbox + section_f_date_of_death_checkbox = 0) OR (section_c_date_of_death_checkbox + section_f_date_of_death_checkbox = 2)) Then err_msg = err_msg & vbCr & "* If you want to update the Date of Death on the MEMB panel, you must select ONLY one date to use for the update."
      
      If (trim(footer_month_updates) = "" OR len(trim(footer_month_updates)) <> 2) OR (trim(footer_year_updates) = "" OR len(trim(footer_year_updates)) <> 2) Then err_msg = err_msg & vbCr & "* If you want to update the STAT panels, you must enter the footer month in the two digit format and the footer year in the two digit format in order for script to make the updates in the correct footer month."
    End If 
    If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    'To do - add parameter for next btn
  LOOP UNTIL err_msg = ""	AND ((ButtonPressed = update_panels_btn) OR (ButtonPressed = skip_panel_updates_btn))								'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If worker indicates that panels should be updated, then script will update panels
'Return to SELF to change the footer month
Call back_to_SELF
EmWriteScreen footer_month_updates, 20, 43
EmWriteScreen footer_year_updates, 20, 46
EmWriteScreen MAXIS_case_number, 18, 43
transmit

If swkr_update_checkbox = 1 Then
  'Ensure the specifically selected assessor (SWKR) will update the panel
  If section_a_assessor_1_checkbox = 1 Then
    swkr_panel_name = section_a_assessor
    swkr_panel_street = section_a_street_address
    swkr_panel_city = section_a_city
    swkr_panel_state = section_a_state
    swkr_panel_zip = section_a_zip_code
    swkr_panel_phone = section_a_phone_number
  ElseIf section_a_assessor_2_checkbox Then
    swkr_panel_name = section_a_assessor_2
    swkr_panel_street = section_a_street_address_2
    swkr_panel_city = section_a_city_2
    swkr_panel_state = section_a_state_2
    swkr_panel_zip = section_a_zip_code_2
    swkr_panel_phone = section_a_phone_number_2
  ElseIf section_a_assessor_3_checkbox Then
    swkr_panel_name = section_a_assessor_3
    swkr_panel_street = section_a_street_address_3
    swkr_panel_city = section_a_city_3
    swkr_panel_state = section_a_state_3
    swkr_panel_zip = section_a_zip_code_3
    swkr_panel_phone = section_a_phone_number_3
  ElseIf section_e_assessor_1_checkbox = 1 Then
    swkr_panel_name = section_e_assessor
    swkr_panel_street = section_e_street_address
    swkr_panel_city = section_e_city
    swkr_panel_state = section_e_state
    swkr_panel_zip = section_e_zip_code
    swkr_panel_phone = section_e_phone_number
  ElseIf section_e_assessor_2_checkbox = 1 Then
    swkr_panel_name = section_e_assessor_2
    swkr_panel_street = section_e_street_address_2
    swkr_panel_city = section_e_city_2
    swkr_panel_state = section_e_state_2
    swkr_panel_zip = section_e_zip_code_2
    swkr_panel_phone = section_e_phone_number_2
  ElseIf section_e_assessor_3_checkbox = 1 Then
    swkr_panel_name = section_e_assessor_3
    swkr_panel_street = section_e_street_address_3
    swkr_panel_city = section_e_city_3
    swkr_panel_state = section_e_state_3
    swkr_panel_zip = section_e_zip_code_3
    swkr_panel_phone = section_e_phone_number_3
  End If

  'Navigate to STAT/SWKR
  Call navigate_to_MAXIS_screen("STAT", "SWKR")
  'Check if SWKR panel exists
  EmReadScreen swkr_does_not_exist, 19, 24, 2
  If swkr_does_not_exist = "SWKR DOES NOT EXIST" Then
    'Add new panel
    Call write_value_and_transmit("NN", 20, 79)
    'Write details to panel
    EMWriteScreen swkr_panel_name, 6, 32
    EMWriteScreen swkr_panel_street, 8, 32
    EMWriteScreen swkr_panel_city, 10, 32
    EMWriteScreen swkr_panel_state, 10, 54
    EMWriteScreen swkr_panel_zip, 10, 63
    EMWriteScreen left(swkr_panel_phone, 3), 12, 34
    EMWriteScreen Mid(swkr_panel_phone, 4, 3), 12, 40
    EMWriteScreen right(swkr_panel_phone, 4), 12, 44
    'Transmit to save 
    transmit
  Else
    'Put panel into edit mode
    PF9
    'Write to panel
    EMWriteScreen swkr_panel_name, 6, 32
    EMWriteScreen swkr_panel_street, 8, 32
    EMWriteScreen swkr_panel_city, 10, 32
    EMWriteScreen swkr_panel_state, 10, 54
    EMWriteScreen swkr_panel_zip, 10, 63
    EMWriteScreen left(swkr_panel_phone, 3), 12, 34
    EMWriteScreen Mid(swkr_panel_phone, 4, 3), 12, 40
    EMWriteScreen right(swkr_panel_phone, 4), 12, 44
    'Transmit to save 
    transmit
  End If
End If

'To do - determine how best to update ADDR panel -> need Living Situation, county of residence code, address line 2
If addr_update_multiple_checkbox = 1 OR addr_update_checkbox_section_c_section_f_match = 1 OR addr_update_checkbox_section_c = 1 OR addr_update_checkbox_section_f = 1 Then
    If addr_update_checkbox_section_c_section_f_match = 1 Then 
      address_to_update = section_f_person_new_address_full
    ElseIf addr_update_multiple_checkbox = 1 Then
      If multiple_section_c_new_address_checkbox = 1 Then
        address_to_update = section_c_person_moved_new_address_full
      ElseIf multiple_section_f_new_address_checkbox = 1 Then
        address_to_update = section_f_person_new_address_full
      End If
    ElseIf addr_update_checkbox_section_c = 1 Then 
      address_to_update = section_c_person_moved_new_address_full
    ElseIf addr_update_checkbox_section_f = 1 Then
      address_to_update = section_f_person_new_address_full
    End If

  BeginDialog Dialog1, 0, 0, 241, 115, "Verify New Address Details"
    Text 5, 5, 225, 10, "Verify the updated address details below to update the ADDR panel:"
    Text 5, 20, 230, 10, address_to_update
    Text 5, 40, 70, 10, "County of Residence:"
    EditBox 80, 35, 60, 15, county_of_residence
    Text 5, 55, 70, 10, "Living Situation:"
    DropListBox 80, 55, 60, 15, "Select one:"+chr(9)+"01"+chr(9)+"02"+chr(9)+"03"+chr(9)+"04"+chr(9)+"05"+chr(9)+"06"+chr(9)+"07"+chr(9)+"08"+chr(9)+"09"+chr(9)+"10", living_situation
    Text 5, 70, 20, 10, "Ver:"
    DropListBox 80, 70, 60, 15, "Select one:"+chr(9)+"SF"+chr(9)+"CO"+chr(9)+"LE"+chr(9)+"MO"+chr(9)+"TX"+chr(9)+"CD"+chr(9)+"UT"+chr(9)+"DL"+chr(9)+"OT"+chr(9)+"NO", address_ver
    ButtonGroup ButtonPressed
      OkButton 130, 95, 50, 15
      CancelButton 185, 95, 50, 15
  EndDialog

  'Dialog validation
  Do
    Do
      'Blank out variables on each new dialog
      err_msg = ""

      dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
      cancel_confirmation

      'Error handling only if worker intends to update panels
      If ButtonPressed = OK Then
        If trim(county_of_residence) = "" OR trim(len(county_of_residence)) <> 2 OR IsNumeric(county_of_residence) = False Then err_msg = err_msg & vbCr & "* The County of Residence field must be filled out with a two-digit number." 
        If living_situation = "Select one:" Then err_msg = err_msg & vbCr & "* You must select an option from the Living Situation dropdown." 
        If address_ver = "Select one:" Then err_msg = err_msg & vbCr & "* You must select an option from the Ver dropdown." 
      End If 
      If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      'To do - add parameter for next btn
    LOOP UNTIL err_msg = ""	AND ButtonPressed = OK								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to STAT/ADDR
  Call navigate_to_MAXIS_screen("STAT", "ADDR")
  ' 'Put panel into edit mode
  ' PF9

  'Write information to panel depending on which address selected
  If addr_update_checkbox_section_c_section_f_match = 1 Then 

    Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

  ElseIf addr_update_multiple_checkbox = 1 Then
    If multiple_section_c_new_address_checkbox = 1 Then

    Call access_ADDR_panel("WRITE", notes_on_address, section_c_street_address, resi_addr_line_two, resi_street_full, section_c_city, section_c_state, section_c_zip_code, county_of_residence, address_ver, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

    ElseIf multiple_section_f_new_address_checkbox = 1 Then

      Call access_ADDR_panel("WRITE", notes_on_address, section_f_person_new_address_address, resi_addr_line_two, resi_street_full, section_f_person_new_address_city, section_f_person_new_address_state, section_f_person_new_address_zip_code, county_of_residence, address_ver, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
    End If

  ElseIf addr_update_checkbox_section_c = 1 Then 

    Call access_ADDR_panel("WRITE", notes_on_address, section_c_street_address, resi_addr_line_two, resi_street_full, section_c_city, section_c_state, section_c_zip_code, county_of_residence, address_ver, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

  ElseIf addr_update_checkbox_section_f = 1 Then

    Call access_ADDR_panel("WRITE", notes_on_address, section_f_person_new_address_address, resi_addr_line_two, resi_street_full, section_f_person_new_address_city, section_f_person_new_address_state, section_f_person_new_address_zip_code, county_of_residence, address_ver, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
  End If
End If

'Update the DOD on MEMB panel
If section_c_person_deceased_checkbox = 1 OR section_f_person_deceased_checkbox = 1 Then
  'Navigate to STAT/MEMB
  Call navigate_to_MAXIS_screen("STAT", "MEMB")
  'Navigate to HH Memb
  Call write_value_and_transmit(left(hh_memb, 2), 20, 76)
  'Put panel into edit mode
  PF9
  'Write date of death ot MEMB panel
  EMReadScreen memb_date_of_death, 10, 19, 42
  If memb_date_of_death = "__ __ ____" then 
    memb_panel_date_of_death_exists = False
  Else
    memb_panel_date_of_death_exists = True
    memb_date_of_death = replace(memb_date_of_death, " ", "/")
  End If
    'If both addresses have been added, then need to compare them to determine if they match
    If section_c_person_deceased_checkbox = 1 AND section_f_person_deceased_checkbox = 1 Then
        'Convert both dates of death to dates to compare them
        section_c_date_of_death = dateadd("m", 0, section_c_date_of_death)
        section_f_person_deceased_date_of_death = dateadd("m", 0, section_f_person_deceased_date_of_death)
        If section_c_date_of_death = section_f_person_deceased_date_of_death Then
            'The dates are the same date
            section_c_section_f_dates_of_death_match = True
        Else
            'The dates do not match - this shouldn't happen
            section_c_section_f_dates_of_death_match = False
        End If
    End If
End If


'THE CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
'Information from DHS 5181 Dialog 1
'Contact information
Call write_variable_in_case_note("~~~DHS-5181 Received~~~")
Call write_variable_in_case_note("Section A - Contact Info")
Call write_bullet_and_variable_in_case_note("Date 5181 sent to worker", section_a_date_form_sent)
Call write_bullet_and_variable_in_case_note("Assessor", section_a_assessor)
Call write_bullet_and_variable_in_case_note("Lead agency", section_a_lead_agency)
Call write_bullet_and_variable_in_case_note("Phone number", section_a_phone_number)
Call write_bullet_and_variable_in_case_note("Address", section_a_street_address & ", " & section_a_city & ", " & section_a_state & ", " & section_a_zip_code)
Call write_bullet_and_variable_in_case_note("Email address", section_a_email_address)
If section_a_assessor_2 <> "" Then
  Call write_variable_in_case_note("Additional Assessor (2)")
  Call write_bullet_and_variable_in_case_note("Assessor", section_a_assessor_2)
  Call write_bullet_and_variable_in_case_note("Lead agency", section_a_lead_agency_2)
  Call write_bullet_and_variable_in_case_note("Phone number", section_a_phone_number_2)
  Call write_bullet_and_variable_in_case_note("Address", section_a_street_address_2 & ", " & section_a_city_2 & ", " & section_a_state_2 & ", " & section_a_zip_code_2)
  Call write_bullet_and_variable_in_case_note("Email address", section_a_email_address_2)
End If
If section_a_assessor_3 <> "" Then
  Call write_variable_in_case_note("Additional Assessor (3)")
  Call write_bullet_and_variable_in_case_note("Assessor", section_a_assessor_3)
  Call write_bullet_and_variable_in_case_note("Lead agency", section_a_lead_agency_3)
  Call write_bullet_and_variable_in_case_note("Phone number", section_a_phone_number_3)
  Call write_bullet_and_variable_in_case_note("Address", section_a_street_address_3 & ", " & section_a_city_3 & ", " & section_a_state_3 & ", " & section_a_zip_code_3)
  Call write_bullet_and_variable_in_case_note("Email address", section_a_email_address_3)
End If
'Person's information
Call write_variable_in_case_note("Person's Information")
Call write_bullet_and_variable_in_case_note("Name", replace(first_name, "_", "") & " " & replace(last_name, "_", "") & "(" & ref_nbr & ")")

'Information from Dialog 2
Call write_variable_in_case_note("Section B - Assessment Results")
Call write_variable_in_case_note("Status")
If section_g_person_requesting_already_enrolled_LTC = 1 Then Call write_bullet_and_variable_in_case_note("Person's current status", "The person currently is requesting services or already enrolled in long-term care services or program")
If section_g_person_will_reside_institution_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Person's current status", "The person resides in or will reside in an institution")
If section_b_diversion_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Program requested", section_b_program_type & " (Diversion)")
ElseIf section_b_conversion_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Program requested", section_b_program_type & " (Conversion)")
Else
  Call write_bullet_and_variable_in_case_note("Program requested", section_b_program_type)
End If
Call write_variable_in_case_note("Institution")
Call write_bullet_and_variable_in_case_note("Admission date", section_b_admission_date)
Call write_bullet_and_variable_in_case_note("Facility", section_b_facility)
Call write_bullet_and_variable_in_case_note("Phone Number", section_b_institution_phone_number)
Call write_bullet_and_variable_in_case_note("Address", section_b_institution_street_address & ", " & section_b_institution_city & ", " & section_b_institution_state & ", " & section_b_institution_zip_code)

'Information from Dialog 3
Call write_variable_in_case_note("Initial Assessment")
Call write_bullet_and_variable_in_case_note("Assessment date", section_b_assessment_date)
Call write_bullet_and_variable_in_case_note("Determination", section_b_assessment_determination)
If section_b_open_to_waiver_yes_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Open to waiver/AC/ECS", "Yes")
If section_b_open_to_waiver_no_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Open to waiver/AC/ECS", "No")
Call write_bullet_and_variable_in_case_note("Estimated monthly waiver/AC costs", section_b_monthly_waiver_costs)
Call write_bullet_and_variable_in_case_note("Anticipated effective date", section_b_waiver_effective_date)
Call write_variable_in_case_note("Case Manager")
If section_b_yes_case_manager = 1 Then Call write_bullet_and_variable_in_case_note("Case manager?", "Yes - I am the case manager")
If section_b_yes_someone_else_case_manager = 1 Then Call write_bullet_and_variable_in_case_note("Case manager?", "Yes - someone else is the manager")
If section_b_no_case_manager = 1 Then 
  Call write_bullet_and_variable_in_case_note("Case manager?", "No")
  Call write_bullet_and_variable_in_case_note("Case manager name", section_b_case_manager_name)
  Call write_bullet_and_variable_in_case_note("Phone number", section_b_case_manager_phone_number)
End If 

'Information from Dialog 4
Call write_variable_in_case_note("Medical Assistance Requests/Applications")
If section_b_applied_MA_LTC_checkbox = 1 Then Call write_variable_in_case_note("Person applied for MA/MA-LTC")
If section_b_ma_enrollee_checkbox = 1 Then 
  Call write_variable_in_case_note("Person is an MA enrollee")
  Call write_bullet_and_variable_in_case_note("Date assessor provided DHS-3543", section_b_date_dhs_3543_provided)
End If
If section_b_completed_dhs_3543_3531_attached_checkbox = 1 Then Call write_variable_in_case_note("Person completed DHS-3543 or DHS-3531 and is attached")
If section_b_completed_dhs_3543_3531_checkbox = 1 Then 
  Call write_variable_in_case_note("Person completed DHS-3543 or DHS-3531 and is attached")
  Call write_bullet_and_variable_in_case_note("Date sent to county", section_b_dhs_3543_3531_sent_to_county_date)
End If
If section_b_send_dhs_3543_checkbox = 1 Then Call write_variable_in_case_note("Send DHS-3543 to person (MA enrollee)")
If section_b_send_dhs_3531_checkbox = 1 Then 
  Call write_variable_in_case_note("Send DHS-3531 to person (not MA enrollee)")
  Call write_bullet_and_variable_in_case_note("Address", section_b_send_dhs_3531_address & ", " & section_b_send_dhs_3531_city & ", " & section_b_send_dhs_3531_state & " " & section_b_send_dhs_3531_zip)
End If
If section_b_send_dhs_3340_checkbox = 1 Then 
  Call write_variable_in_case_note("Send DHS-3340 to person (asset assessment needed)")
  Call write_bullet_and_variable_in_case_note("Address", section_b_send_dhs_3340_address & ", " & section_b_send_dhs_3340_city & ", " & section_b_send_dhs_3340_state & " " & section_b_send_dhs_3340_zip)
End If
Call write_variable_in_case_note("Changes completed by assessor at reassessment")
If section_b_person_no_longer_institutional_LOC_checkbox = 1 Then 
  Call write_variable_in_case_note("Person no longer meets institutional LOC")
  Call write_bullet_and_variable_in_case_note("Effective date of waiver exit no sooner than", section_b_date_waiver_exit)
End If
If section_b_person_enroll_another_program = 1 Then 
  Call write_bullet_and_variable_in_case_note("Person chooses to enroll in another program", section_b_enroll_another_program_list)
End If

'Information from Dialog 5
Call write_variable_in_case_note("Exit Reasons")
If section_c_exited_waiver_program_checkbox = 1 Then 
  Call write_variable_in_case_note("Person exited waiver program")
  Call write_bullet_and_variable_in_case_note("Effective date of waiver exit", section_c_date_waiver_exit)
End If
Call write_variable_in_case_note("   Reasons for exit")
'Create the list of exit reasons as a variable
exit_reasons = ""
If section_c_hospital_admission_checkbox = 1 Then exit_reasons = exit_reasons & "hospital admission, "
If section_c_nursing_facility_admission_checkbox = 1 Then exit_reasons = exit_reasons & "nursing facility admission, "
If section_c_person_informed_choice_checkbox = 1 Then exit_reasons = exit_reasons & "person's informed choice, "
If section_c_person_deceased_checkbox = 1 Then exit_reasons = exit_reasons & "person is deceased" & " (DOD: " & section_c_date_of_death & ")"
If section_c_person_moved_out_of_state_checkbox = 1 Then exit_reasons = exit_reasons & "person moved out of state" & " (date of move: " & section_c_date_of_move & ")"
If section_c_exited_for_other_reasons_checkbox = 1 Then exit_reasons = exit_reasons & "exited for other reasons: " & section_c_exited_for_other_reasons_explanation

'Information from Dialog 6
Call write_variable_in_case_note("Other Changes")
If section_c_diversion_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Program requested", section_c_program_type & " (Diversion)")
ElseIf section_c_conversion_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Program requested", section_c_program_type & " (Conversion)")
Else
  Call write_bullet_and_variable_in_case_note("Program requested", section_c_program_type)
End If
If section_c_person_moved_new_address_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Date address changed", section_c_date_address_changed)
  Call write_bullet_and_variable_in_case_note("Address", section_c_street_address & ", " & section_c_city & ", " & section_c_state & " " & section_c_zip_code)
End If
If section_c_new_legal_rep_checkbox = 1 Then
  Call write_bullet_and_variable_in_case_note("Person has new legal representative", section_c_legal_rep_first_name & " " & section_c_legal_rep_last_name & "(" & section_c_legal_rep_phone_number & ")")
  Call write_bullet_and_variable_in_case_note("Legal rep address", section_c_legal_rep_street_address & ", " & section_c_legal_rep_city & ", " & section_c_legal_rep_state & " " & section_c_legal_rep_zip_code)
End If
If section_c_person_return_to_community_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Person returning to community w/in 121 days of qual admission", "Effective date: " & section_c_qual_admission_eff_date)
If section_c_other_changes_program_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Other changes related to program/service elig", section_c_other_changes_program)
Call write_variable_in_case_note("Comments - Assessor, case manager, or care coordinator")
If section_d_additional_comments = 1 Then Call write_bullet_and_variable_in_case_note("Additional notes or comments", section_d_additional_comments)

'Information from Dialog 7
'Contact information
Call write_variable_in_case_note("Section E")
Call write_bullet_and_variable_in_case_note("Date 5181 sent to worker", section_e_date_form_sent)
Call write_bullet_and_variable_in_case_note("Assessor", section_e_assessor)
Call write_bullet_and_variable_in_case_note("Lead agency", section_e_lead_agency)
Call write_bullet_and_variable_in_case_note("Phone number", section_e_phone_number)
Call write_bullet_and_variable_in_case_note("Address", section_e_street_address & ", " & section_e_city & ", " & section_e_state & ", " & section_e_zip_code)
Call write_bullet_and_variable_in_case_note("Email address", section_e_email_address)
If section_e_assessor_2 <> "" Then
  Call write_variable_in_case_note("Additional Assessor (2)")
  Call write_bullet_and_variable_in_case_note("Assessor", section_e_assessor_2)
  Call write_bullet_and_variable_in_case_note("Lead agency", section_e_lead_agency_2)
  Call write_bullet_and_variable_in_case_note("Phone number", section_e_phone_number_2)
  Call write_bullet_and_variable_in_case_note("Address", section_e_street_address_2 & ", " & section_e_city_2 & ", " & section_e_state_2 & ", " & section_e_zip_code_2)
  Call write_bullet_and_variable_in_case_note("Email address", section_e_email_address_2)
End If
If section_e_assessor_3 <> "" Then
  Call write_variable_in_case_note("Additional Assessor (3)")
  Call write_bullet_and_variable_in_case_note("Assessor", section_e_assessor_3)
  Call write_bullet_and_variable_in_case_note("Lead agency", section_e_lead_agency_3)
  Call write_bullet_and_variable_in_case_note("Phone number", section_e_phone_number_3)
  Call write_bullet_and_variable_in_case_note("Address", section_e_street_address_3 & ", " & section_e_city_3 & ", " & section_e_state_3 & ", " & section_e_zip_code_3)
  Call write_bullet_and_variable_in_case_note("Email address", section_e_email_address_3)
End If

'Information from Dialog 8
Call write_variable_in_case_note("Section F - Medical Assistance")
Call write_variable_in_case_note("MA status for long-term supports and services")
If section_f_person_applied_MA_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Person applied for MA/MA-LTC", section_f_person_applied_date)
If section_f_dhs_3531_sent_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("DHS-3531 sent to person", section_f_dhs_3531_sent_date)
If section_f_dhs_3543_sent_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("DHS-3543 sent to person", section_f_dhs_3543_sent_date)
If section_f_dhs_3543_3531_returned_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("DHS-3543/3531 returned; elig. determ. pending", section_f_dhs_3543_3531_returned_comments)
If section_f_dhs_3543_3531_not_returned_checkbox = 1 Then Call write_variable_in_case_note("DHS-3543/DHS-3531 has not been returned")
Call write_variable_in_case_note("Determination")
If section_f_ma_opened_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("MA opened (effective date)", section_f_ma_opened_date)
If section_f_basic_ma_medical_spenddown_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Basic MA medical spenddown", section_f_basic_ma_medical_spenddown)
If section_f_ma_LTC_services_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("MA for LTC services open on specific date", section_f_ma_LTC_services_date)
If section_f_LTC_spenddown_initial_month_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("LTC spenddown/waiver olbig. for initial month", section_f_LTC_spenddown_date)
If section_f_ma_denied_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("MA denied", section_f_ma_denied_date)
If section_f_ma_payment_denied_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("MA payment of LTC services denied", section_f_ma_payment_LTC_date)
If section_f_inelig_for_MA_payment_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Person inelig for MA payment of LTSS services until specific date", section_f_inelig_for_MA_payment_date)
If section_f_basic_ma_continues_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Basic MA continues until specific date", section_f_basic_ma_continues_date)
If section_f_asset_assessment_results_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Results from asset assessment sent to person", section_f_results_from_asset_assessment_sent_date)

'Information from Dialog 9
Call write_variable_in_case_note("Changes")
If section_f_LTC_spenddown_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("LTC spenddown/waiver obligation", section_f_LTC_spenddown_amount)
If section_f_MA_terminated_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("MA terminated - basic MA & payment of LTSS services", section_f_ma_terminated_eff_date)
If section_f_basic_ma_spenddown_change_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Basic MA spenddown changed", section_f_basic_ma_spenddown_change_amount)
If section_f_ma_payment_terminated_basic_open_checkbox = 1 Then 
  Call write_bullet_and_variable_in_case_note("Date terminated", section_f_ma_payment_terminated_term_date)
  Call write_bullet_and_variable_in_case_note("Date inelig. through", section_f_ma_payment_terminated_date_inelig_thru)
End If
If section_f_person_deceased_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Person is deceased", section_f_person_deceased_date_of_death)
If section_f_person_moved_institution_checkbox = 1 Then
  Call write_variable_in_case_note("Person moved to an institution") 
  Call write_bullet_and_variable_in_case_note("Date of admission", section_f_person_moved_institution_admit_date)
  Call write_bullet_and_variable_in_case_note("Facility name and phone number", section_f_person_moved_institution_facility_name & "(" & section_f_person_moved_institution_phone_number & ")")
  Call write_bullet_and_variable_in_case_note("Address", section_f_person_moved_institution_address & ", " & section_f_person_moved_institution_city & ", " & section_f_person_moved_institution_state & " " & section_f_person_moved_institution_zip)
End If
If section_f_person_new_address_checkbox = 1 Then 
  Call write_variable_in_case_note("Person has a new address")
  Call write_bullet_and_variable_in_case_note("Date of address change", section_f_person_new_address_date_changed)
  If section_f_person_new_address_new_phone_number <> "" Then Call write_bullet_and_variable_in_case_note("New phone number", section_f_person_new_address_new_phone_number)
  Call write_bullet_and_variable_in_case_note("Address", section_f_person_new_address_address & ", " & section_f_person_new_address_city & ", " & section_f_person_new_address_state & " " & section_f_person_new_address_zip_code)
End If
If section_f_other_change_checkbox = 1 Then Call write_bullet_and_variable_in_case_note("Other change", section_f_person_other_change_description)

'Info from Dialog 10
Call write_variable_in_case_note("Section G: Comments from eligibility worker")
If section_g_elig_comments <> "" Then Call write_bullet_and_variable_in_case_note("Additional comments", section_g_elig_comments)

'Add worker signature
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