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
contact_info_btn                = 201
status_btn                      = 202
initial_assessment_btn          = 203
MA_req_app_btn                  = 204
exit_reasons_btn                = 205
other_changes_btn               = 206
section_d_comments_btn          = 207
contact_info_btn                = 208  
MA_status_determination_btn     = 209
changes_btn                     = 210  
section_g_comments_btn          = 211
next_btn                        = 212
previous_btn                    = 213
complete_btn                    = 214

'--Other buttons
instructions_btn                        = 215
section_a_add_assessor_btn              = 216
section_e_add_assessor_btn              = 217
section_a_assessor_return_btn           = 218
section_e_assessor_return_btn           = 219
section_a_assessor_return_no_save_btn   = 220
section_e_assessor_return_no_save_btn   = 221

'Defining variables
dialog_count = ""

'DEFINING FUNCTIONS===========================================================================

'Dialog 1 - Section A: Contact Information
function section_a_contact_info()
  dialog_count = 1
  BeginDialog Dialog1, 0, 0, 326, 310, "1 - Section A: Contact Information"
  GroupBox 5, 5, 245, 175, "FROM (assessor/case manager/care coordinator's information)"
  Text 10, 25, 70, 10, "Date Sent to Worker:"
  Text 10, 40, 40, 10, "Assessor:"
  Text 10, 55, 50, 10, "Lead Agency:"
  Text 10, 70, 55, 10, "Phone Number:"
  Text 10, 85, 55, 10, "Street Address:"
  Text 10, 100, 20, 10, "City:"
  Text 10, 115, 25, 10, "State:"
  Text 10, 130, 35, 10, "Zip Code:"
  Text 10, 145, 55, 10, "Email Address:"
  Text 10, 165, 145, 10, "Click button to add up to 2 add'l assessors:"
  EditBox 90, 20, 55, 15, section_a_date_form_sent
  EditBox 90, 35, 150, 15, section_a_assessor
  EditBox 90, 50, 150, 15, section_a_lead_agency
  EditBox 90, 65, 55, 15, section_a_phone_number
  EditBox 90, 80, 150, 15, section_a_street_address
  EditBox 90, 95, 150, 15, section_a_city
  EditBox 90, 110, 25, 15, section_a_state
  EditBox 90, 125, 55, 15, section_a_zip_code
  EditBox 90, 140, 150, 15, section_a_email_address
  ButtonGroup ButtonPressed
   PushButton 160, 160, 85, 15, "Add/Update Assessor", section_a_add_assessor_btn
  GroupBox 5, 190, 245, 30, "Person's Information"
  Text 10, 200, 70, 10, "Select HH Member:"
  DropListBox 80, 200, 160, 15, HH_Memb_DropDown, hh_memb
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
EndDialog
end function
'To do - dim all the variables?
Dim section_a_date_form_sent, section_a_assessor, section_a_lead_agency, section_a_phone_number, section_a_street_address, section_a_city, section_a_state, section_a_zip_code, section_a_email_address, hh_memb

function section_a_additional_assessors()
  dialog_count = 11
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
      PushButton 205, 290, 120, 15, "Save Info and Return to Contact Info", section_a_assessor_return_btn
      PushButton 5, 290, 185, 15, "Return to Contact Info WITHOUT Saving Assessor Info", section_a_assessor_return_no_save_btn
  EndDialog
end function
'Dim all variables in function
Dim section_a_assessor_2, section_a_lead_agency_2, section_a_phone_number_2, section_a_street_address_2, section_a_city_2, section_a_state_2, section_a_zip_code_2, section_a_email_address_2, section_a_assessor_3, section_a_lead_agency_3, section_a_phone_number_3, section_a_street_address_3, section_a_city_3, section_a_state_3, section_a_zip_code_3, section_a_email_address_3

'Dialog 2 - Section B: Assessment Results - Current Status
function section_b_assess_results_current_status()
  dialog_count = 2
  BeginDialog Dialog1, 0, 0, 326, 310, "2 - Section B: Assess. Results - Current Status"
  GroupBox 5, 5, 250, 50, "What is the person's current status? (check second if both apply)"
  CheckBox 15, 20, 10, 10, "", section_g_person_requesting_already_enrolled_LTC
  Text 25, 20, 215, 20, "The person currently is requesting services or already enrolled in long-term care services or program"
  CheckBox 15, 40, 195, 10, "The person resides in or will reside in an institution", section_g_person_will_reside_institution_checkbox
  GroupBox 5, 60, 250, 55, "Program Type"
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
      PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
      PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
  BeginDialog Dialog1, 0, 0, 326, 310, "6 - Section C: Other Changes & Section D: Comments"
  GroupBox 5, 5, 250, 235, "Other changes"
  Text 15, 20, 50, 10, "Program type"
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
    PushButton 265, 235, 50, 15, "Changes", changes_btn
    PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
EndDialog
end function
'Dim the variables in function
Dim section_c_program_type_list, section_c_diversion_checkbox, section_c_conversion_checkbox, section_c_person_moved_new_address_checkbox, section_c_date_address_changed, section_c_street_address, section_c_city, section_c_state, section_c_zip_code, section_c_new_legal_rep_checkbox, section_c_legal_rep_first_name, section_c_legal_rep_last_name, section_c_legal_rep_phone_number, section_c_legal_rep_street_address, section_c_legal_rep_city, section_c_legal_rep_state, section_c_legal_rep_zip_code, section_c_person_return_to_community_checkbox, section_c_qual_admission_eff_date, section_c_other_changes_program_checkbox, section_c_other_changes_program, section_d_additional_comments

'Dialog 7 - Section E: Contact Information
function section_e_contact_info()
  dialog_count = 7
  first_name
  last_name
  BeginDialog Dialog1, 0, 0, 326, 310, "7 - Section E: Contact Information"
    Text 10, 10, 105, 10, "Date Sent to Eligibility Worker:"
    EditBox 120, 5, 75, 15, section_e_date_form_sent
    GroupBox 5, 25, 245, 155, "TO (assessor/case manager/care coordinator's information)"
    Text 10, 40, 40, 10, "Assessor:"
    EditBox 90, 35, 150, 15, section_e_assessor
    Text 10, 55, 50, 10, "Lead Agency:"
    EditBox 90, 50, 150, 15, section_e_lead_agency
    Text 10, 70, 55, 10, "Phone Number:"
    EditBox 90, 65, 55, 15, section_e_phone_number
    Text 10, 85, 55, 10, "Street Address:"
    EditBox 90, 80, 150, 15, section_e_street_address
    Text 10, 100, 20, 10, "City:"
    EditBox 90, 95, 150, 15, section_e_city
    Text 10, 115, 25, 10, "State:"
    EditBox 90, 110, 25, 15, section_e_state
    Text 10, 130, 35, 10, "Zip Code:"
    EditBox 90, 125, 45, 15, section_e_zip_code
    Text 10, 145, 55, 10, "Email Address:"
    EditBox 90, 140, 150, 15, section_e_email_address
    Text 10, 165, 145, 10, "Click button to add up to 2 add'l assessors:"
    ButtonGroup ButtonPressed
      PushButton 160, 160, 85, 15, "Add/Update Assessor", section_e_add_assessor_btn
    GroupBox 5, 190, 245, 80, "Person's Information"
    Text 10, 205, 105, 10, "Information entered previously:"
    Text 15, 215, 40, 10, "First name:"
    Text 70, 215, 170, 10, first_name
    Text 15, 225, 40, 10, "Last name:"
    Text 70, 225, 170, 10, last_name
    Text 15, 235, 45, 10, "Ref Number:"
    Text 70, 235, 75, 10, ref_nbr
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
      PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
      PushButton 205, 290, 120, 15, "Save Info and Return to Contact Info", section_e_assessor_return_btn
      PushButton 5, 290, 185, 15, "Return to Contact Info WITHOUT Saving Assessor Info", section_e_assessor_return_no_save_btn
  EndDialog
end function
' Dim all functions in variable
Dim section_e_assessor_2, section_e_lead_agency_2, section_e_phone_number_2, section_e_street_address_2, section_e_city_2, section_e_state_2, section_e_zip_code_2, section_e_email_address_2, section_e_assessor_3,  section_e_lead_agency_3, section_e_phone_number_3, section_e_street_address_3, section_e_city_3, section_e_state_3, section_e_zip_code_3, section_e_email_address_3

'Dialog 8 - Section F: Medical Assistance
function section_f_medical_assistance()
  dialog_count = 8
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
      PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
      PushButton 265, 55, 50, 15, "Status", status_btn
      PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
      PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
      PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
      PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
      PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
      PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
      PushButton 265, 220, 50, 15, "MA Status/Det", MA_status_determination_btn
      PushButton 265, 235, 50, 15, "Changes", changes_btn
      PushButton 265, 265, 50, 15, "Comments", section_g_comments_btn
  EndDialog
end function
' Dim all functions in variable
Dim section_f_LTC_spenddown_checkbox, section_f_LTC_spenddown_amount, section_f_MA_terminated_checkbox,section_f_ma_terminated_eff_date, section_f_basic_ma_spenddown_change_checkbox, section_f_basic_ma_spenddown_change_amount, section_f_ma_payment_terminated_basic_open_checkbox, section_f_ma_payment_terminated_term_date, section_f_ma_payment_terminated_date_inelig_thru, section_f_person_deceased_date_of_death, section_f_person_moved_institution_checkbox, section_f_person_moved_institution_admit_date, section_f_person_moved_institution_facility_name, section_f_person_moved_institution_phone_number, section_f_person_moved_institution_address, section_f_person_moved_institution_city, section_f_person_moved_institution_state, section_f_person_moved_institution_zip, section_f_person_new_address_checkbox, section_f_person_new_address_date_changed, section_f_person_new_address_new_phone_number, section_f_person_new_address_address, section_f_person_new_address_city, section_f_person_new_address_state, section_f_person_new_address_zip_code, section_f_other_change_checkbox, section_f_person_other_change_description

'Dialog 10 - Section G: Comments from eligibility worker
function section_g_comments_elig_worker()
  dialog_count = 10
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
    PushButton 265, 25, 50, 15, "Contact Info", contact_info_btn
    PushButton 265, 55, 50, 15, "Status", status_btn
    PushButton 265, 70, 50, 15, "Initial Assess.", initial_assessment_btn
    PushButton 265, 85, 50, 15, "MA Req/App", MA_req_app_btn
    PushButton 265, 115, 50, 15, "Exit Reasons", exit_reasons_btn
    PushButton 265, 130, 50, 15, "Other Changes", other_changes_btn
    PushButton 265, 160, 50, 15, "Comments", section_d_comments_btn
    PushButton 265, 190, 50, 15, "Contact Info", contact_info_btn
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
	If ButtonPressed = next_btn AND err_msg = "" Then dialog_count = dialog_count + 1 'If next is selected, it will go to the next dialog
	If ButtonPressed = previous_btn AND err_msg = "" Then dialog_count = dialog_count - 1	'If previous is selected, it will go to the previous dialog

  If err_msg = "" and ButtonPressed = contact_info_btn then dialog_count = 1
  If err_msg = "" and ButtonPressed = status_btn then dialog_count = 2
  If err_msg = "" and ButtonPressed = initial_assessment_btn then dialog_count = 3
  If err_msg = "" and ButtonPressed = MA_req_app_btn then dialog_count = 4
  If err_msg = "" and ButtonPressed = exit_reasons_btn then dialog_count = 5
  If err_msg = "" and ButtonPressed = other_changes_btn then dialog_count = 6
  If err_msg = "" and ButtonPressed = section_d_comments_btn then dialog_count = 6
  If err_msg = "" and ButtonPressed = contact_info_btn then dialog_count = 7
  If err_msg = "" and ButtonPressed = MA_status_determination_btn then dialog_count = 8
  If err_msg = "" and ButtonPressed = changes_btn then dialog_count = 9
  If err_msg = "" and ButtonPressed = section_g_comments_btn then dialog_count = 10
  If err_msg = "" and ButtonPressed = section_a_add_assessor_btn then dialog_count = 11
  If err_msg = "" and ButtonPressed = section_e_add_assessor_btn then dialog_count = 12
  If err_msg = "" and ButtonPressed = section_a_assessor_return_btn then dialog_count = 1
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
  If err_msg = "" and ButtonPressed = section_e_assessor_return_btn then dialog_count = 7
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
end function

function dialog_specific_error_handling()	'Error handling for main dialog of forms
  'Error handling will display at the point of each dialog and will not let the user continue unless the applicable errors are resolved. Had to list all buttons including -1 so ensure the error reporting is called and hit when the script is run.
  'To do - need these?
	If ButtonPressed = contact_info_btn OR _
    ButtonPressed = status_btn OR _
    ButtonPressed = initial_assessment_btn OR _
    ButtonPressed = MA_req_app_btn OR _
    ButtonPressed = exit_reasons_btn OR _
    ButtonPressed = other_changes_btn OR _
    ButtonPressed = section_d_comments_btn OR _
    ButtonPressed = contact_info_btn OR _  
    ButtonPressed = MA_status_determination_btn OR _
    ButtonPressed = changes_btn OR _  
    ButtonPressed = section_g_comments_btn OR _
    ButtonPressed = next_btn OR _
    ButtonPressed = previous_btn OR _
    ButtonPressed = complete_btn OR _
    ButtonPressed = instructions_btn OR _
    ButtonPressed = section_a_assessor_return_btn OR _
    ButtonPressed = section_e_assessor_return_btn OR _
    ButtonPressed = -1 Then
      If dialog_count = 1 then 
        If trim(section_a_date_form_sent) = "" OR IsDate(section_a_date_form_sent) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY." 
        If trim(section_a_assessor) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor field." 
        If trim(section_a_lead_agency) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency field." 
        If trim(section_a_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field." 
        If trim(section_a_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field." 
        If trim(section_a_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field." 
        If trim(section_a_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State field." 
        If trim(section_a_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field." 
        If trim(section_a_email_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address field." 
        If hh_memb = "Select One:" Then err_msg = err_msg & vbNewLine & "* You must select the Household Member from the dropdown." 
      End If
      If dialog_count = 11 then 
        If trim(section_a_assessor_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (2) field." 
        If trim(section_a_lead_agency_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (2) field." 
        If trim(section_a_phone_number_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field." 
        If trim(section_a_street_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (2) field." 
        If trim(section_a_city_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (2) field." 
        If trim(section_a_state_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State (2) field." 
        If trim(section_a_zip_code_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (2) field." 
        If trim(section_a_email_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (2) field." 

        'Handling for Asessor (3) to only trigger errors if some fields are filled in but if completely blank then it will ignore errors
        If trim(section_a_assessor_3) <> "" or trim(section_a_lead_agency_3) <> "" OR trim(section_a_phone_number_3) <> "" OR trim(section_a_street_address_3) <> "" OR trim(section_a_city_3) <> "" OR trim(section_a_state_3) <> "" OR trim(section_a_zip_code_3) <> "" OR trim(section_a_email_address_3) <> "" Then
          If trim(section_a_assessor_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (3) field." 
          If trim(section_a_lead_agency_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (3) field." 
          If trim(section_a_phone_number_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field." 
          If trim(section_a_street_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (3) field." 
          If trim(section_a_city_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (3) field." 
          If trim(section_a_state_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State (3) field." 
          If trim(section_a_zip_code_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (3) field." 
          If trim(section_a_email_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (3) field." 
        End If
      End If
      If dialog_count = 12 then 
        If trim(section_e_assessor_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (2) field." 
        If trim(section_e_lead_agency_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (2) field." 
        If trim(section_e_phone_number_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (2) field." 
        If trim(section_e_street_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (2) field." 
        If trim(section_e_city_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (2) field." 
        If trim(section_e_state_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State (2) field." 
        If trim(section_e_zip_code_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code (2) field." 
        If trim(section_e_email_address_2) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address (2) field." 

        'Handling for Asessor (3) to only trigger errors if some fields are filled in but if completely blank then it will ignore errors
        If trim(section_e_assessor_3) <> "" or trim(section_e_lead_agency_3) <> "" OR trim(section_e_phone_number_3) <> "" OR trim(section_e_street_address_3) <> "" OR trim(section_e_city_3) <> "" OR trim(section_e_state_3) <> "" OR trim(section_e_zip_code_3) <> "" OR trim(section_e_email_address_3) <> "" Then
          If trim(section_e_assessor_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor (3) field." 
          If trim(section_e_lead_agency_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency (3) field." 
          If trim(section_e_phone_number_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number (3) field." 
          If trim(section_e_street_address_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address (3) field." 
          If trim(section_e_city_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City (3) field." 
          If trim(section_e_state_3) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State (3) field." 
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
        If trim(section_b_institution_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field."
        If trim(section_b_institution_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field."
        If trim(section_b_institution_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field."
        If trim(section_b_institution_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State field."
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
          If trim(section_b_case_manager_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field for the case manager."
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
          If trim(section_c_street_address) = "" OR trim(section_c_city) = "" or trim(section_c_state) OR trim(section_c_zip_code) Then err_msg = err_msg & vbNewLine & "* You must fill out the fields for the new address (address, state, city, and zip code)."
        End If
        If section_c_new_legal_rep_checkbox = 1 Then
          If trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the First Name, Last Name, and Phone Number fields for the new legal representative."
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
        If section_c_program_type_list = "Select one:" Then err_msg = err_msg & vbNewLine & "* You must select the program type from the dropdown list."

        If section_c_diversion_checkbox + section_c_conversion_checkbox = 2 Then err_msg = err_msg & vbNewLine & "* You can only select one option, not both, for Diversion or Conversion."
        
        If section_c_person_moved_new_address_checkbox = 1 Then
          If trim(section_c_date_address_changed) = "" or IsDate(section_c_date_address_changed) = False Then err_msg = err_msg & vbNewLine & "* You must enter the Date of Address Change in the format MM/DD/YYYY."
          If trim(section_c_street_address) = "" OR trim(section_c_city) = "" or trim(section_c_state) OR trim(section_c_zip_code) Then err_msg = err_msg & vbNewLine & "* You must fill out the fields for the new address (address, state, city, and zip code)."
        End If
        If section_c_new_legal_rep_checkbox = 1 Then
          If trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_first_name) = "" or trim(section_c_legal_rep_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the First Name, Last Name, and Phone Number fields for the new legal representative."
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
      If dialog_count = 8 then 
        If trim(section_e_date_form_sent) = "" OR IsDate(section_e_date_form_sent) = FALSE Then err_msg = err_msg & vbNewLine & "* You must fill out the Date Sent to Worker field in the format MM/DD/YYYY." 
        If trim(section_e_assessor) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Assessor field." 
        If trim(section_e_lead_agency) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Lead Agency field." 
        If trim(section_e_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Phone Number field." 
        If trim(section_e_street_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Street Address field." 
        If trim(section_e_city) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the City field." 
        If trim(section_e_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the State field." 
        If trim(section_e_state) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Zip Code field." 
        If trim(section_e_email_address) = "" Then err_msg = err_msg & vbNewLine & "* You must fill out the Email Address field." 
        If hh_memb = "Select One:" Then err_msg = err_msg & vbNewLine & "* You must select the Household Member from the dropdown." 
      End If
      If dialog_count = 9 then
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
      If dialog_count = 10 then
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
          If trim(section_f_person_moved_institution_admit_date) = "" OR IsDate(section_f_person_moved_institution_admit_date) Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date in the format MM/DD/YYYY."
          If trim(section_f_person_moved_institution_facility_name) = "" OR trim(section_f_person_moved_institution_phone_number) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date, facility name, and phone number for the institution."
          If trim(section_f_person_moved_institution_address) = "" OR trim(section_f_person_moved_institution_city) = "" OR trim(section_f_person_moved_institution_state) = "" OR trim(section_f_person_moved_institution_zip) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the address, city, state, and zip code for the institution."
        End If

        If section_f_person_new_address_checkbox = 1 Then
          If trim(section_f_person_new_address_date_changed) = "" OR IsDate(section_f_person_new_address_date_changed) Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date in the format MM/DD/YYYY."
          If trim(section_f_person_new_address_new_phone_number) = "" OR trim(section_f_person_new_address_address) = "" OR trim(section_f_person_new_address_city) = "" OR trim(section_f_person_new_address_state) = "" OR trim(section_f_person_new_address_zip_code) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the admit date, facility name, and phone number for the institution."
          If trim(section_f_person_moved_institution_address) = "" OR trim(section_f_person_moved_institution_city) = "" OR trim(section_f_person_moved_institution_state) = "" OR trim(section_f_person_moved_institution_zip) = "" Then err_msg = err_msg & vbNewLine & "* You checked the box indicating that the person moved to an institution. You must enter the address, city, state, and zip code for the institution."
        End If

        If section_f_other_change_checkbox = 1 Then
          If trim(section_f_person_other_change_description) = "" Then err_msg = err_msg & vbNewLine & "* You checked the Other change box. You must describe the reason in the field provided."
        End If
      End If
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
			'To do - add form specific handling
      Call dialog_specific_error_handling	'function for error handling of main dialog of forms
			Call button_movement()				'function to move throughout the dialogs
		Loop until err_msg = ""
	Loop until ButtonPressed = complete_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

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