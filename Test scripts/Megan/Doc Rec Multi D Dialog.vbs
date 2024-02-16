'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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
EMConnect "" 'Connects to BlueZone


'DEFINING CONSTANTS & ARRAY===========================================================================
'Define Constants
const form_type_const   = 0
const btn_name_const    = 1
const btn_number_const	= 2
const count_of_form		= 3
const the_last_const	= 4

'Defining array capturing form names, button names, button numbers
Dim form_type_array()		'Defining 1D array
ReDim form_type_array(the_last_const, 0)	'Redefining array so we can resize it 
form_count = 0				'Counter for array should start with 0
all_form_array = "*"
false_count = 0 


'Dim/ReDim Array for form checkbox selections
Dim unchecked, checked		'Defining unchecked/checked 
unchecked = 0			
checked = 1


'Defining count of forms
asset_count 	= 0 
atr_count 		= 0 
arep_count 		= 0 
change_count	= 0 
evf_count		= 0 
hosp_count		= 0 
iaa_count		= 0 
iaa_ssi_count	= 0
ltc_1503_count	= 0 
mof_count		= 0 
mtaf_count		= 0 
psn_count		= 0 
sf_count		= 0 
diet_count		= 0


'Button Defined
add_button 			= 201
all_forms 			= 202
'review_selections 	= 203
clear_button		= 204
next_btn			= 205
previous_btn		= 206
complete_btn		= 207

asset_btn			= 400
atr_btn				= 401
arep_btn			= 402
change_btn 			= 403
evf_btn				= 404
hospice_btn			= 405
iaa_btn				= 406
iaa_ssi_btn			= 407
ltc_1503_btn		= 408
mof_btn				= 409
mtaf_btn			= 410
psn_btn				= 411
sf_btn				= 412
diet_btn			= 413

'Define resource buttons
iaa_CM121203_btn			= 2000	
iaa_sp_pben_btn				= 2001
iaa_te021214_btn			= 2002
iaa_ssi_CM121203_btn		= 2003
iaa_ssi_sp_pben_btn			= 2004
iaa_ssi_te021214_btn		= 2005
diet_link_CM_special_diet	= 2006
diet_SP_referrals			= 2007
hosp_TE0207081_btn			= 2008
hosp_SP_hospice_btn			= 2009
psn_CM1315_btn				= 2010
psn_TE1817_btn				= 2011
psn_hss_btn					= 2012
psn_mhm_btn					= 2013
psn_hsss_btn				= 2014


'Define Form Names
asset_form_name 	= "Asset Statement"
atr_form_name		= "Authorization to Release Information (ATR)"
arep_form_name		= "AREP (Authorized Rep)"
change_form_name	= "Change Report Form"
evf_form_name		= "Employment Verification Form (EVF)"
hosp_form_name		= "Hospice Transaction Form"
iaa_form_name		= "Interim Assistance Agreement (IAA)"
iaa_ssi_form_name	= "Interim Assistance Authorization- SSI"
ltc_1503_form_name	= "LTC-1503"
mof_form_name		= "Medical Opinion Form (MOF)"
mtaf_form_name		= "Minnesota Transition Application Form (MTAF)"
psn_form_name		= "Professional Statement of Need (PSN)"
sf_form_name		= "Residence and Shelter Expenses Release Form"
diet_form_name		= "Special Diet Information Request"

'FUNCTIONS DEFINED===========================================================================
function asset_dialog()
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, asset_effective_date
			EditBox 310, 20, 45, 15, asset_date_received
			EditBox 30, 65, 270, 15, asset_Q1
			EditBox 30, 85, 270, 15, asset_Q2
			EditBox 30, 105, 270, 15, asset_Q3
			EditBox 30, 125, 270, 15, asset_Q4
			Text 5, 5, 220, 10, asset_form_name
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Responses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""
end function

function atr_dialog()
	EditBox 175, 15, 45, 15, atr_effective_date
	EditBox 300, 15, 45, 15, atr_date_received		
	DropListBox 50, 40, 100, 15, HH_Memb_DropDown, atr_member_dropdown
	EditBox 205, 40, 45, 15, atr_start_date
	EditBox 300, 40, 45, 15, atr_end_date
	DropListBox 80, 60, 70, 15, ""+chr(9)+"Verbal"+chr(9)+"Written", atr_authorization_type
	DropListBox 65, 95, 60, 15, ""+chr(9)+"Organization"+chr(9)+"Person", atr_contact_type
	EditBox 160, 95, 170, 15, atr_name
	EditBox 70, 115, 175, 15, atr_address
	EditBox 35, 135, 85, 15, atr_city
	DropListBox 155, 135, 30, 15, ""+chr(9)+"AL"+chr(9)+"AK"+chr(9)+"AZ"+chr(9)+"AR"+chr(9)+"CA"+chr(9)+"CO"+chr(9)+"CT"+chr(9)+"DE"+chr(9)+"DC"+chr(9)+"FL"+chr(9)+"GA"+chr(9)+"HI"+chr(9)+"ID"+chr(9)+"IL"+chr(9)+"IN"+chr(9)+"IA"+chr(9)+"KS"+chr(9)+"KY"+chr(9)+"LA"+chr(9)+"ME"+chr(9)+"MD"+chr(9)+"MA"+chr(9)+"MI"+chr(9)+"MN"+chr(9)+"MS"+chr(9)+"MO"+chr(9)+"MT"+chr(9)+"NE"+chr(9)+"NV"+chr(9)+"NH"+chr(9)+"NJ"+chr(9)+"NM"+chr(9)+"NY"+chr(9)+"NC"+chr(9)+"ND"+chr(9)+"OH"+chr(9)+"OK"+chr(9)+"OR"+chr(9)+"PA"+chr(9)+"RI"+chr(9)+"SC"+chr(9)+"SD"+chr(9)+"TN"+chr(9)+"TX"+chr(9)+"UT"+chr(9)+"VT"+chr(9)+"VA"+chr(9)+"WA"+chr(9)+"WV"+chr(9)+"WI"+chr(9)+"WY", atr_state
	EditBox 230, 135, 35, 15, atr_zipcode
	EditBox 70, 160, 75, 15, atr_phone_number
	CheckBox 35, 200, 170, 10, "to continue evaluation or treatment", atr_eval_treat_checkbox
	CheckBox 35, 210, 170, 10, "to coordinate services", atr_coor_serv_checkbox
	CheckBox 35, 220, 170, 10, "to determine eligibility for assistance/service", atr_elig_serv_checkbox
	CheckBox 35, 230, 170, 10, "for court proceedings", atr_court_checkbox
	CheckBox 35, 240, 80, 10, "other (specify below)", atr_other_checkbox
	EditBox 50, 250, 90, 15, atr_other
	EditBox 50, 280, 230, 15, atr_comments
	Text 5, 5, 220, 10, "ATR- AUTHORIZATION TO RELEASE"
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, MAXIS_case_number
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 55, 10, "Document Date:"
	Text 15, 45, 30, 10, "Member"
	Text 170, 45, 35, 10, "Start Date"
	Text 265, 45, 30, 10, "End Date"
	Text 15, 65, 65, 10, "Authorization Type"
	GroupBox 10, 85, 340, 95, "Contact Person/Organization"
	Text 20, 100, 45, 10, "Contact Type"
	Text 140, 100, 20, 10, "Name"
	Text 20, 120, 50, 10, "Street Address"
	Text 20, 140, 15, 10, "City"
	Text 135, 140, 20, 10, "State"
	Text 200, 140, 30, 10, "Zip code"
	Text 20, 165, 50, 10, "Phone Number"
	GroupBox 10, 190, 335, 80, "Record requested will be used: "
	Text 10, 285, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function
Dim atr_effective_date, atr_date_received, atr_member_dropdown, atr_start_date, atr_end_date, atr_authorization_type, atr_contact_type, atr_name, atr_address, atr_city, atr_state, atr_zipcode, atr_phone_number, atr_eval_treat_checkbox, atr_coor_serv_checkbox, atr_elig_serv_checkbox, atr_court_checkbox, atr_other_checkbox, atr_other, atr_comments

function arep_dialog()
	Text 60, 25, 45, 10, MAXIS_case_number
	EditBox 175, 20, 45, 15, arep_effective_date
	EditBox 310, 20, 45, 15, arep_date_received		
	EditBox 45, 55, 185, 15, arep_name
	EditBox 45, 75, 185, 15, arep_street
	EditBox 45, 95, 85, 15, arep_city
	EditBox 160, 95, 20, 15, arep_state
	EditBox 200, 95, 30, 15, arep_zip
	EditBox 45, 115, 50, 15, arep_phone_one
	EditBox 115, 115, 20, 15, arep_ext_one
	EditBox 45, 135, 50, 15, arep_phone_two
	EditBox 115, 135, 20, 15, arep_ext_two
	CheckBox 20, 160, 60, 10, "Forms to AREP", arep_forms_to_arep_checkbox
	CheckBox 95, 160, 75, 10, "MMIS Mail to AREP", arep_mmis_mail_to_arep_checkbox
	CheckBox 20, 175, 185, 10, "Check here to have the script update the AREP Panel", arep_update_AREP_panel_checkbox
	EditBox 120, 200, 50, 15, AREP_recvd_date
	CheckBox 20, 220, 75, 10, "ID on file for AREP?", AREP_ID_check
	CheckBox 20, 235, 215, 10, "TIKL to get new HC form 12 months after date form was signed?", arep_TIKL_check
	EditBox 140, 250, 65, 15, arep_signature_date
	CheckBox 255, 55, 115, 10, "AREP Req - MHCP - DHS-3437", arep_dhs_3437_checkbox
	CheckBox 255, 75, 105, 10, "AREP Req - HC12729", arep_HC_12729_checkbox
	CheckBox 255, 95, 100, 10, "SNAP AREP Choice - D405", arep_D405_checkbox
	CheckBox 255, 115, 105, 10, "AREP on CAF", arep_CAF_AREP_page_checkbox
	CheckBox 255, 135, 100, 10, "AREP on any HC App", arep_HCAPP_AREP_checkbox
	CheckBox 255, 155, 75, 10, "Power of Attorney", arep_power_of_attorney_checkbox
	CheckBox 265, 215, 35, 10, "SNAP", arep_SNAP_AREP_checkbox
	CheckBox 265, 225, 50, 10, "Health Care", arep_HC_AREP_checkbox
	CheckBox 265, 235, 30, 10, "Cash", arep_CASH_AREP_checkbox
	Text 5, 5, 220, 10, arep_form_name
	Text 5, 25, 50, 10, "Case Number:"
	Text 125, 25, 50, 10, "Effective Date:"
	Text 245, 25, 60, 10, "Document Date:"
	GroupBox 10, 45, 225, 145, "Panel Information"
	Text 20, 60, 25, 10, "Name:"
	Text 20, 80, 25, 10, "Street:"
	Text 25, 100, 15, 10, "City:"
	Text 140, 100, 20, 10, "State:"
	Text 185, 100, 15, 10, "Zip:"
	Text 20, 120, 25, 10, "Phone:"
	Text 100, 120, 15, 10, "Ext."
	Text 20, 140, 25, 10, "Phone:"
	Text 100, 140, 15, 10, "Ext."
	Text 20, 205, 95, 10, "Date of AREP Form Received"
	Text 20, 255, 115, 10, "Date form was signed (MM/DD/YY):"
	GroupBox 245, 45, 125, 155, "Specific FORM Received"
	Text 275, 65, 50, 10, "(HC)"
	Text 275, 85, 60, 10, "(Cash and SNAP)"
	Text 275, 105, 75, 10, "(SNAP and EBT Card)"
	Text 275, 125, 60, 10, "(Cash and SNAP)"
	Text 275, 145, 50, 10, "(HC)"
	Text 275, 165, 60, 10, "(HC, SNAP, Cash)"
	Text 250, 175, 110, 20, "Checking the FORM will indicate the programs in the CASE/NOTE"
	Text 255, 205, 85, 10, "Programs Authorized for:"
	Text 395, 35, 45, 10, "    --Forms--"
end function
Dim arep_effective_date, arep_date_received, arep_name, arep_street, arep_city, arep_state, arep_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, arep_forms_to_arep_checkbox, arep_mmis_mail_to_arep_checkbox, arep_update_AREP_panel_checkbox, AREP_recvd_date, AREP_ID_check, arep_TIKL_check, arep_signature_date, arep_dhs_3437_checkbox, arep_HC_12729_checkbox, arep_D405_checkbox, arep_CAF_AREP_page_checkbox, arep_HCAPP_AREP_checkbox, arep_power_of_attorney_checkbox, arep_SNAP_AREP_checkbox, arep_HC_AREP_checkbox, arep_CASH_AREP_checkbox

function change_dialog()
	EditBox 175, 15, 45, 15, chng_effective_date
	EditBox 310, 15, 45, 15, chng_date_received		
	EditBox 50, 45, 320, 15, chng_address_notes
	EditBox 50, 65, 320, 15, chng_household_notes
	EditBox 110, 85, 260, 15, chng_asset_notes
	EditBox 50, 105, 320, 15, chng_vehicles_notes
	EditBox 50, 125, 320, 15, chng_income_notes
	EditBox 50, 145, 320, 15, chng_shelter_notes
	EditBox 50, 165, 320, 15, chng_other_change_notes
	EditBox 65, 200, 305, 15, chng_actions_taken
	EditBox 65, 220, 305, 15, chng_other_notes
	EditBox 75, 240, 295, 15, chng_verifs_requested
	DropListBox 105, 265, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", chng_notable_change
	CheckBox 10, 285, 140, 10, "Check here to navigate to DAIL/WRIT", chng_tikl_nav_check		'TODO: handling for tikl - nav to dail/writ
	DropListBox 270, 280, 95, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", chng_changes_continue
	Text 5, 5, 220, 10, change_form_name
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, MAXIS_case_number
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	GroupBox 5, 35, 370, 150, "CHANGES REPORTED"
	Text 15, 50, 30, 10, "Address:"
	Text 15, 70, 35, 10, "HH Comp:"
	Text 15, 90, 95, 10, "Assets (savings or property):"
	Text 15, 110, 30, 10, "Vehicles:"
	Text 15, 130, 30, 10, "Income:"
	Text 15, 150, 25, 10, "Shelter:"
	Text 15, 170, 20, 10, "Other:"
	GroupBox 5, 190, 370, 70, "ACTIONS"
	Text 15, 205, 45, 10, "Action Taken:"
	Text 15, 225, 45, 10, "Other Notes:"
	Text 15, 245, 60, 10, "Verifs Requested:"
	Text 10, 270, 95, 10, "Notable changes reported? "
	Text 180, 285, 90, 10, "The changes client reports:"
	Text 395, 35, 45, 10, "    --Forms--"

end function
Dim chng_effective_date, chng_date_received, chng_address_notes, chng_household_notes, chng_asset_notes, chng_vehicles_notes, chng_income_notes, chng_shelter_notes, chng_other_change_notes, chng_actions_taken, chng_other_notes, chng_verifs_requested, chng_changes_continue, chng_notable_change 'Change Reported variables


function evf_dialog()
	Text 60, 25, 45, 10, MAXIS_case_number
	EditBox 175, 20, 45, 15, evf_effective_date
	EditBox 310, 20, 45, 15, evf_date_received		
	ComboBox 95, 50, 210, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
	EditBox 95, 70, 210, 15, evf_employer
	DropListBox 95, 90, 210, 15, HH_Memb_DropDown, evf_client
	DropListBox 100, 145, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", evf_info
	EditBox 245, 145, 60, 15, evf_info_date
	EditBox 100, 165, 60, 15, evf_request_info
	CheckBox 185, 170, 105, 10, "Create TIKL for additional info", EVF_TIKL_checkbox
	EditBox 95, 205, 210, 15, evf_actions_taken
	Text 5, 5, 220, 10, evf_form_name
	Text 5, 25, 50, 10, "Case Number:"
	Text 125, 25, 50, 10, "Effective Date:"
	Text 245, 25, 60, 10, "Document Date:"
	Text 50, 55, 40, 10, "EVF Status:"
	Text 35, 75, 55, 10, "Employer name:"
	Text 30, 95, 60, 10, "Household Memb:"
	GroupBox 30, 130, 285, 60, "Is additional information needed?"
	Text 35, 150, 60, 10, "Addt'l Info Reqstd:"
	Text 185, 150, 55, 10, "Date Requested:"
	Text 35, 170, 65, 10, "Info Requested via:"
	Text 40, 210, 50, 10, "Actions taken:"
	Text 395, 35, 45, 10, "    --Forms--"
end function 
Dim evf_effective_date, evf_date_received, EVF_status_dropdown, evf_employer, evf_client, evf_info, evf_info_date, evf_request_info, EVF_TIKL_checkbox, evf_actions_taken

function hospice_dialog()
	EditBox 175, 20, 45, 15, hosp_effective_date
	EditBox 310, 20, 45, 15, hosp_date_received		
	DropListBox 100, 45, 165, 15, HH_Memb_DropDown, hosp_resident_name
	EditBox 100, 65, 205, 15, hosp_name
	EditBox 100, 85, 80, 15, hosp_npi_number
	EditBox 100, 105, 50, 15, hosp_entry_date
	EditBox 205, 105, 50, 15, hosp_exit_date
	EditBox 100, 125, 50, 15, hosp_mmis_updated_date
	EditBox 30, 160, 275, 15, hosp_reason_not_updated
	EditBox 30, 190, 275, 15, hosp_other_notes
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "TE 02.07.081", hosp_TE0207081_btn
		PushButton 65, 280, 50, 15, "MA-Hospice", hosp_SP_hospice_btn
	Text 5, 5, 220, 10, hosp_form_name
	Text 5, 25, 50, 10, "Case Number:"
	Text 125, 25, 50, 10, "Effective Date:"
	Text 245, 25, 60, 10, "Document Date:"
	Text 50, 50, 45, 10, "Client Name:"
	Text 35, 70, 60, 10, "Name of Hospice:"
	Text 50, 90, 45, 10, "NPI Number:"
	Text 55, 110, 40, 10, "Entry Date:"
	Text 170, 110, 35, 10, "Exit Date:"
	Text 30, 130, 70, 10, "MMIS Updated as of "
	Text 30, 150, 165, 10, "If MMIS has not yet been updated, explain reason:"
	Text 30, 180, 50, 10, "Other Notes:"
	Text 60, 25, 45, 10, MAXIS_case_number
	Text 395, 35, 45, 10, "    --Forms--"		
end function 
Dim hosp_effective_date, hosp_date_received, hosp_resident_name, hosp_name, hosp_npi_number, hosp_entry_date, hosp_exit_date, hosp_mmis_updated_date, hosp_reason_not_updated, hosp_other_notes, hosp_TE0207081_btn, hosp_SP_hospice_btn

function iaa_dialog()
	Text 60, 20, 45, 10, MAXIS_case_number
	EditBox 175, 15, 45, 15, iaa_effective_date
	EditBox 310, 15, 45, 15, iaa_date_received		
	DropListBox 55, 45, 140, 15, HH_Memb_DropDown, iaa_member_dropdown
	DropListBox 305, 45, 60, 15, ""+chr(9)+"Initial claim"+chr(9)+"Post-eligibility", iaa_type_assistance
	EditBox 50, 80, 145, 15, iaa_benefits_1
	EditBox 220, 80, 145, 15, iaa_benefits_3
	EditBox 50, 100, 145, 15, iaa_benefits_2
	EditBox 220, 100, 145, 15, iaa_benefits_4
	DropListBox 95, 140, 115, 15, ""+chr(9)+"01-RSDI"+chr(9)+"02-SSI"+chr(9)+"06-Child Support"+chr(9)+"07-Alimony"+chr(9)+"08-VA Disability"+chr(9)+"09-VA Pension"+chr(9)+"10-VA Dependent Educational"+chr(9)+"11-VA Dependent Other"+chr(9)+"12-Unemployment Insurance"+chr(9)+"13-Worker's Comp"+chr(9)+"14-RR Retirement"+chr(9)+"15-Other Ret"+chr(9)+"16-Military Allot"+chr(9)+"17-EITC"+chr(9)+"18-Strike Pay"+chr(9)+"19-Other"+chr(9)+"21-SMRT", iaa_benefit_type
	EditBox 310, 140, 55, 15, iaa_referral_date
	DropListBox 95, 160, 115, 15, ""+chr(9)+"1-Copy of Chkstb"+chr(9)+"2-Award Letters"+chr(9)+"4-Coltrl Stmt"+chr(9)+"5-Other Document"+chr(9)+"N-No Ver Prvd", iaa_verification_dropdown
	EditBox 310, 160, 55, 15, iaa_date_applied_pben
	DropListBox 95, 180, 115, 15, ""+chr(9)+"A-Appealing"+chr(9)+"D-Denied"+chr(9)+"E-Eligible"+chr(9)+"P-Pending"+chr(9)+"N-Not Appl Yet"+chr(9)+"R-Refused To Accept", iaa_disposition_code_dropdown
	EditBox 310, 180, 55, 15, iaa_date
	CheckBox 20, 215, 310, 15, "NOT signed within 30 days of receiving CAF or Change Report Form. Casenote invalid IAA.", iaa_not_signed_30_checkbox
	CheckBox 20, 235, 225, 10, "Check here to have the script update PBEN", iaa_update_pben_checkbox
	EditBox 50, 255, 315, 15, iaa_comments
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "CM12.12.03", iaa_CM121203_btn
		PushButton 65, 280, 50, 15, "TE02.12.14", iaa_te021214_btn
		PushButton 125, 280, 70, 15, "HSRM- PBEN Panel", iaa_sp_pben_btn
	Text 5, 5, 220, 10, "INTERIM ASSISTANCE AUTHORIZATION"
	Text 5, 20, 50, 10, "Case Number:"
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 20, 50, 30, 10, "Member"
	Text 215, 50, 90, 10, "Type of interim assistance"
	Text 20, 70, 130, 10, "Other benefits you may be eligible for"
	GroupBox 15, 125, 360, 75, "PBEN Fields"
	Text 50, 145, 45, 10, "Benefit Type"
	Text 265, 145, 45, 10, "Referral Date"
	Text 50, 165, 40, 10, "Verification"
	Text 235, 165, 75, 10, "Date Applied for PBEN"
	Text 35, 185, 60, 10, "Disposition Code"
	Text 280, 185, 30, 10, "IAA Date"
	GroupBox 15, 205, 360, 40, "Actions"
	Text 15, 260, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function 
Dim iaa_effective_date, iaa_date_received, iaa_member_dropdown, iaa_type_assistance, iaa_benefits_1, iaa_benefits_2, iaa_benefits_3, iaa_benefits_4, iaa_benefit_type, iaa_referral_date, iaa_verification_dropdown, iaa_date_applied_pben, iaa_disposition_code_dropdown, iaa_date, iaa_not_signed_30_checkbox, iaa_update_pben_checkbox, iaa_comments


function iaa_ssi_dialog()
	Text 60, 20, 45, 10, MAXIS_case_number
	EditBox 175, 15, 45, 15, iaa_ssi_effective_date
	EditBox 310, 15, 45, 15, iaa_ssi_date_received		
	DropListBox 50, 40, 120, 15, HH_Memb_DropDown, iaa_ssi_member_dropdown 
	DropListBox 275, 40, 95, 20, ""+chr(9)+"General Assistance (GA)"+chr(9)+"Housing Support (HS)", iaa_ssi_type_of_interim_assistance
	DropListBox 100, 75, 120, 15, ""+chr(9)+"01-RSDI"+chr(9)+"02-SSI"+chr(9)+"06-Child Support"+chr(9)+"07-Alimony"+chr(9)+"08-VA Disability"+chr(9)+"09-VA Pension"+chr(9)+"10-VA Dependent Educational"+chr(9)+"11-VA Dependent Other"+chr(9)+"12-Unemployment Insurance"+chr(9)+"13-Worker's Comp"+chr(9)+"14-RR Retirement"+chr(9)+"15-Other Ret"+chr(9)+"16-Military Allot"+chr(9)+"17-EITC"+chr(9)+"18-Strike Pay"+chr(9)+"19-Other"+chr(9)+"21-SMRT", iaa_ssi_benefit_type
	EditBox 315, 75, 55, 15, iaa_ssi_referral_date
	DropListBox 100, 95, 120, 15, ""+chr(9)+"1-Copy of Chkstb"+chr(9)+"2-Award Letters"+chr(9)+"4-Coltrl Stmt"+chr(9)+"5-Other Document"+chr(9)+"N-No Ver Prvd", iaa_ssi_verification_dropdown
	EditBox 315, 95, 55, 15, iaa_ssi_date_applied_pben
	DropListBox 100, 115, 120, 15, ""+chr(9)+"A-Appealing"+chr(9)+"D-Denied"+chr(9)+"E-Eligible"+chr(9)+"P-Pending"+chr(9)+"N-Not Appl Yet"+chr(9)+"R-Refused To Accept", iaa_ssi_disposition_code_dropdown
	EditBox 315, 115, 55, 15, iaa_ssi_iaa_date
	CheckBox 30, 155, 225, 10, "Check here to have the script update PBEN", iaa_ssi_update_pben_checkbox
	CheckBox 30, 165, 310, 15, "NOT signed within 30 days of receiving CAF or Change Report Form. Casenote invalid IAA.", iaa_ssi_not_signed_30days_checkbox
	EditBox 55, 195, 315, 15, iaa_ssi_comments
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "CM12.12.03", iaa_ssi_CM121203_btn
		PushButton 65, 280, 50, 15, "TE02.12.14", iaa_ssi_te021214_btn
		PushButton 125, 280, 75, 15, "HSRM- PBEN Panel", iaa_ssi_sp_pben_btn
	Text 5, 5, 220, 10, iaa_ssi_form_name
	Text 5, 20, 50, 10, "Case Number:"
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 15, 45, 30, 10, "Member"
	Text 185, 45, 90, 10, "Type of interim assistance"
	GroupBox 20, 60, 360, 75, "PBEN Fields"
	Text 55, 80, 45, 10, "Benefit Type"
	Text 270, 80, 45, 10, "Referral Date"
	Text 55, 100, 40, 10, "Verification"
	Text 240, 100, 75, 10, "Date Applied for PBEN"
	Text 40, 120, 60, 10, "Disposition Code"
	Text 285, 120, 30, 10, "IAA Date"
	GroupBox 20, 145, 360, 40, "Actions"
	Text 20, 200, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function
Dim iaa_ssi_effective_date, iaa_ssi_date_received, iaa_ssi_member_dropdown, iaa_ssi_type_of_interim_assistance, iaa_ssi_benefit_type, iaa_ssi_referral_date, iaa_ssi_verification_dropdown, iaa_ssi_date_applied_pben, iaa_ssi_disposition_code_dropdown, iaa_ssi_iaa_date, iaa_ssi_update_pben_checkbox, iaa_ssi_not_signed_30days_checkbox, iaa_ssi_comments, iaa_ssi_CM121203_btn, iaa_ssi_te021214_btn, iaa_ssi_sp_pben_btn

function ltc_1503_dialog()
	Text 60, 25, 45, 10, MAXIS_case_number
	EditBox 175, 20, 45, 15, ltc_1503_effective_date
	EditBox 310, 20, 45, 15, ltc_1503_date_received
	EditBox 60, 50, 110, 15, ltc_1503_FACI_1503
	DropListBox 230, 50, 70, 15, ""+chr(9)+"30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", ltc_1503_length_of_stay
	DropListBox 105, 70, 40, 15, ""+chr(9)+"SNF"+chr(9)+"NF"+chr(9)+"ICF-DD"+chr(9)+"RTC", ltc_1503_level_of_care
	DropListBox 205, 70, 80, 15, ""+chr(9)+"acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-DD", ltc_1503_admitted_from
	EditBox 60, 90, 90, 15, ltc_1503_hospital_admitted_from
	EditBox 210, 90, 45, 15, ltc_1503_admit_date
	EditBox 320, 90, 45, 15, ltc_1503_discharge_date
	CheckBox 15, 110, 155, 10, "If you've processed this 1503, check here.", ltc_1503_processed_1503_checkbox
	CheckBox 10, 145, 65, 10, "Updated RLVA?", ltc_1503_updated_RLVA_checkbox
	CheckBox 80, 145, 60, 10, "Updated FACI?", ltc_1503_updated_FACI_checkbox
	CheckBox 145, 145, 50, 10, "Need 3543?", ltc_1503_need_3543_checkbox
	CheckBox 205, 145, 55, 10, "Need 3531?", ltc_1503_need_3531_checkbox
	CheckBox 265, 145, 95, 10, "Need asset assessment?", ltc_1503_need_asset_assessment_checkbox
	EditBox 125, 155, 210, 15, ltc_1503_verifs_needed
	CheckBox 10, 180, 85, 10, "Sent 3050 back to LTCF", ltc_1503_sent_3050_checkbox
	CheckBox 105, 180, 70, 10, "Sent verif req? To:", ltc_1503_sent_verif_request_checkbox
	ComboBox 175, 175, 60, 15, ""+chr(9)+"client"+chr(9)+"AREP"+chr(9)+"Client & AREP", ltc_1503_sent_request_to
	CheckBox 245, 180, 120, 10, "Sent DHS-5181 to Case Manager", ltc_1503_sent_5181_checkbox
	CheckBox 10, 210, 255, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", ltc_1503_TIKL_checkbox
	CheckBox 10, 225, 155, 10, "Check here to have the script update HCMI.", ltc_1503_HCMI_update_checkbox
	CheckBox 10, 240, 150, 10, "Check here to have the script update FACI.", ltc_1503_FACI_update_checkbox
	EditBox 100, 260, 25, 15, ltc_1503_faci_footer_month
	EditBox 130, 260, 25, 15, ltc_1503_faci_footer_year
	EditBox 245, 260, 75, 15, ltc_1503_mets_case_number
	EditBox 30, 280, 330, 15, ltc_1503_notes
	Text 5, 5, 220, 10, "LTC 1503	"
	Text 5, 25, 50, 10, "Case Number:"
	Text 125, 25, 50, 10, "Effective Date:"
	Text 245, 25, 60, 10, "Document Date:"
	GroupBox 0, 40, 370, 85, "Facility Info"
	Text 10, 55, 50, 10, "Facility name:"
	Text 175, 55, 50, 10, "Length of stay:"
	Text 10, 75, 95, 10, "Recommended level of care:"
	Text 155, 75, 50, 10, "Admitted from:"
	Text 10, 95, 50, 10, "Hospital Name"
	Text 155, 95, 55, 10, "Admission Date:"
	Text 265, 95, 55, 10, "Discharge Date:"
	GroupBox 5, 135, 365, 60, "Actions/Proofs"
	Text 10, 160, 115, 10, "Other proofs needed (if applicable):"
	GroupBox 0, 200, 370, 55, "Script actions"
	Text 5, 265, 95, 10, "Facility Update Month/Year:"
	Text 165, 265, 75, 10, "METS Case Number:"
	Text 5, 285, 25, 10, "Notes:"
	Text 395, 35, 45, 10, "    --Forms--"
end function
Dim ltc_1503_effective_date, ltc_1503_date_received, ltc_1503_FACI_1503, ltc_1503_length_of_stay, ltc_1503_level_of_care, ltc_1503_admitted_from, ltc_1503_hospital_admitted_from, ltc_1503_admit_date, ltc_1503_discharge_date, ltc_1503_processed_1503_checkbox, ltc_1503_updated_RLVA_checkbox, ltc_1503_updated_FACI_checkbox, ltc_1503_need_3543_checkbox, ltc_1503_need_3531_checkbox, ltc_1503_need_asset_assessment_checkbox, ltc_1503_verifs_needed, ltc_1503_sent_3050_checkbox, ltc_1503_sent_verif_request_checkbox, ltc_1503_sent_request_to, ltc_1503_sent_5181_checkbox, ltc_1503_TIKL_checkbox, ltc_1503_HCMI_update_checkbox, ltc_1503_FACI_update_checkbox, ltc_1503_faci_footer_month, ltc_1503_faci_footer_year, ltc_1503_mets_case_number, ltc_1503_notes

function mof_dialog()
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, mof_effective_date
			EditBox 310, 20, 45, 15, mof_date_received
			EditBox 30, 65, 270, 15, mof_Q1
			EditBox 30, 85, 270, 15, mof_Q2
			EditBox 30, 105, 270, 15, mof_Q3
			EditBox 30, 125, 270, 15, mof_Q4			
			Text 5, 5, 220, 10, mof_form_name
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Responses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""
end function 

function mtaf_dialog()	
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, mtaf_effective_date
			EditBox 310, 20, 45, 15, mtaf_date_received
			EditBox 30, 65, 270, 15, mtaf_Q1
			EditBox 30, 85, 270, 15, mtaf_Q2
			EditBox 30, 105, 270, 15, mtaf_Q3
			EditBox 30, 125, 270, 15, mtaf_Q4			
			Text 5, 5, 220, 10, mtaf_form_name
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Responses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""
end function

function psn_dialog()
	Text 60, 20, 45, 10, MAXIS_case_number
	EditBox 180, 15, 45, 15, psn_date_received	
	DropListBox 270, 15, 100, 15, HH_Memb_DropDown, psn_member_dropdown
	DropListBox 15, 45, 105, 15, ""+CHR(9)+"Yes- At least 1 selected"+chr(9)+"No- Section NOT completed", psn_section_1_dropdown
	DropListBox 15, 60, 105, 15, ""+CHR(9)+"Yes- 1 selected"+chr(9)+"No- Section NOT completed", psn_section_2_dropdown
	DropListBox 15, 75, 105, 15, ""+CHR(9)+"Yes- At least 1 selected"+chr(9)+"No- Section NOT completed", psn_section_3_dropdown
	DropListBox 15, 90, 105, 15, ""+CHR(9)+"Yes- At least 2 selected"+chr(9)+"No- Section NOT completed", psn_section_4_dropdown
	DropListBox 15, 105, 105, 15, ""+CHR(9)+"Yes- Section completed"+chr(9)+"No- Section NOT completed", psn_section_5_dropdown
	EditBox 95, 120, 120, 15, psn_cert_prof
	EditBox 250, 120, 125, 15, psn_facility	'TODO: PUll vendor # & name from FACI panel and auto fill. User can edit if necessary. 
	DropListBox 65, 155, 30, 15, ""+CHR(9)+"Y"+chr(9)+"N", psn_wreg_fs_pwe
	DropListBox 195, 155, 155, 15, ""+CHR(9)+"03-Unfit for Employment"+chr(9)+"04-Resp for Care of Incapacitated Person"+chr(9)+"05-Age 60 or Older"+chr(9)+"06-Under Age 16"+chr(9)+"07-Age 16-17, Living w/ Caregiver"+chr(9)+"08-Resp for Care of Child under 6"+chr(9)+"09-Empl 30 hrs/wk or Earnings of 30 hrs/wk"+chr(9)+"10-Matching Grant Participant"+chr(9)+"11-Receiving or Applied for UI"+chr(9)+"12-Enrolled in School, Training, or Higher Ed"+chr(9)+"13-Participating in CD Program"+chr(9)+"14-Receiving MFIP"+chr(9)+"20-Pending/Receiving DWP"+chr(9)+ "15-Age 16-17, NOT Living w/ Caregiver"+chr(9)+"16-50-59 Years Old"+chr(9)+"17-Receiving RCA or GA"+chr(9)+"21-Resp for Care of Child under 18"+chr(9)+"23-Pegnant", psn_wreg_work_wreg_status
	DropListBox 65, 175, 115, 15, ""+CHR(9)+"01-Work Reg Exempt"+chr(9)+"02-Under Age 18"+chr(9)+"03-Age 50 or Over"+chr(9)+"04-Caregiver of Minor Child"+chr(9)+"05-Pregnant"+chr(9)+"06-Employed Avg of 20 hrs/wk"+chr(9)+"07-Work Experience Participant"+chr(9)+"08-ther E&T Services"+chr(9)+"09-Resides in a Waivered Area"+chr(9)+"10-ABAWD Counted Month"+chr(9)+"11-2nd-3rd Month Period of Elig"+chr(9)+"12-RCA or GA Recipient"+chr(9)+"13-ABAWD Banked Months", psn_wreg_abawd_status
	DropListBox 255, 175, 130, 20, ""+CHR(9)+"04-Permanent Ill or Incap"+chr(9)+"05-Temporary Ill or Incap"+chr(9)+"06-Care of Ill or Incap Mbr"+chr(9)+"07-Requires Services In Residence"+chr(9)+"09-Mntl Ill or Dev Disabled"+chr(9)+"10-SSI/RSDI Pend"+chr(9)+"11-Appealing SSI/RSDI Denial"+chr(9)+"12-Advanced Age"+chr(9)+"13-Learning Disability"+chr(9)+"17-Protect/Court Ordered"+chr(9)+"20-Age 16 or 17 SS Approval "+chr(9)+"25-Emancipated Minor"+chr(9)+"28-Unemployable"+chr(9)+"29-Displaced Hmkr (Ft Student)"+chr(9)+"30-Minor w/ Adult Unrelated"+chr(9)+"32-ESL, Adult/HS At least half time"+chr(9)+"35-Drug/Alcohol Addiction (DAA)"+chr(9)+"99-No Elig Basis", psn_wreg_ga_elig_status
	EditBox 65, 195, 45, 15, psn_disa_begin_date
	EditBox 255, 195, 45, 15, psn_disa_end_date
	EditBox 65, 215, 45, 15, psn_disa_cert_start
	EditBox 255, 215, 45, 15, psn_disa_cert_end
	DropListBox 65, 235, 110, 15, ""+CHR(9)+"01-RSDI Only Disability"+chr(9)+"02-RSDI Only Blindness"+chr(9)+"03-SSI, SSI/RSDI Disability"+chr(9)+"04-SSI, SSI/RSDI Blindness"+chr(9)+"06-SMRT/SSA Pend"+chr(9)+"08-SMRT Certified Blindness"+chr(9)+"09-Ill/Incapacity"+chr(9)+"10-SMRT Certified Disability", psn_disa_status
	DropListBox 255, 235, 105, 15, ""+CHR(9)+"1-DHS161/Dr Stmt"+chr(9)+"2-SMRT Certified"+chr(9)+"3-Certified For RSDI or SSI"+chr(9)+"6-Other Document"+chr(9)+"7-Professional Stmt of Need"+chr(9)+"N-No Ver Prvd", psn_disa_verif
	EditBox 55, 260, 320, 15, psn_comments
	ButtonGroup ButtonPressed
		PushButton 10, 280, 40, 15, "CM 13.15", psn_CM1315_btn
		PushButton 55, 280, 35, 15, "TE18.17", psn_TE1817_btn
		PushButton 95, 280, 30, 15, "HSS", psn_hss_btn
		PushButton 130, 280, 30, 15, "MHM", psn_mhm_btn
		PushButton 165, 280, 30, 15, "HSSS", psn_hsss_btn
	Text 5, 5, 130, 10, psn_form_name
	Text 5, 20, 50, 10, "Case Number:"
	Text 125, 20, 55, 10, "Document Date:"
	Text 235, 20, 30, 10, "Member"
	GroupBox 5, 35, 385, 115, "PSN Fields"
	Text 125, 50, 105, 10, "Section 1: Housing Situation"
	Text 125, 65, 105, 10, "Section 2: Disabling Condition"
	Text 125, 80, 150, 10, "Section 3: MA Housing Stabilization Services"
	Text 125, 95, 185, 10, "Section 4: MN Housing Support Supplemental Services"
	Text 125, 110, 220, 10, "Section 5: Transition from Residential Treatment to MN HS Program"
	Text 20, 125, 72, 10, "Certified Professional:"
	Text 225, 125, 25, 10, "Facility:"
	GroupBox 5, 145, 385, 110, "WREG/DISA Panel Updates"
	Text 30, 160, 30, 10, "FS PWE:"
	Text 115, 160, 80, 10, "FSET Work Reg Status: "
	Text 10, 180, 55, 10, "ABAWD Status: "
	Text 190, 180, 65, 10, "GA Elig Basis Code:"
	Text 10, 200, 55, 10, "Disa Begin Date: "
	Text 205, 200, 50, 10, "Disa End Date:"
	Text 25, 220, 40, 10, "Cert Period:"
	Text 205, 220, 50, 10, "Cert End Date:"
	Text 395, 35, 45, 10, "    --Forms--"
	Text 25, 240, 40, 10, "Disa Status: "
	Text 215, 240, 40, 10, "Verification:"
	Text 15, 265, 37, 10, "Comments:"
	Text 395, 35, 45, 10, "    --Forms--"
end function 
Dim  psn_date_received, psn_member_dropdown, psn_section_1_dropdown, psn_section_2_dropdown, psn_section_3_dropdown, psn_section_4_dropdown, psn_section_5_dropdown, psn_cert_prof, psn_facility, psn_wreg_fs_pwe, psn_wreg_work_wreg_status, psn_wreg_abawd_status, psn_wreg_ga_elig_status, psn_disa_begin_date, psn_disa_end_date, psn_disa_cert_start, psn_disa_cert_end, psn_disa_status, psn_disa_verif, psn_comments

function sf_dialog()	
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, sf_effective_date
			EditBox 310, 20, 45, 15, sf_date_received
			EditBox 30, 65, 270, 15, sf_Q1
			EditBox 30, 85, 270, 15, sf_Q2
			EditBox 30, 105, 270, 15, sf_Q3
			EditBox 30, 125, 270, 15, sf_Q4			
			Text 5, 5, 220, 10, sf_form_name
			Text 125, 25, 50, 10, "Effective Date:"
			Text 15, 70, 10, 10, "Q1"
			Text 245, 25, 60, 10, "Document Date:"
			GroupBox 5, 50, 305, 195, "Responses to form questions captured here"
			Text 5, 25, 50, 10, "Case Number:"
			Text 395, 35, 45, 10, "    --Forms--"
			Text 15, 110, 10, 10, "Q3"
			Text 15, 130, 15, 10, "Q4"
			Text 15, 90, 15, 10, "Q2"
			Text 15, 150, 15, 10, ""
end function

function diet_dialog()
	EditBox 175, 15, 45, 15, diet_effective_date
	EditBox 310, 15, 45, 15, diet_date_received		
	DropListBox 50, 35, 120, 15, HH_Memb_DropDown, diet_member_number 
	DropListBox 55, 70, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_1_dropdown
	DropListBox 185, 70, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_1_dropdown
	DropListBox 290, 70, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_1_dropdown
	DropListBox 55, 85, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_2_dropdown
	DropListBox 185, 85, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_2_dropdown
 	DropListBox 290, 85, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_2_dropdown
	DropListBox 55, 100, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_3_dropdown
	DropListBox 185, 100, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_3_dropdown
	DropListBox 290, 100, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_3_dropdown
	DropListBox 55, 115, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_4_dropdown
	DropListBox 185, 115, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_4_dropdown
	DropListBox 290, 115, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_4_dropdown
	DropListBox 55, 130, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_5_dropdown
	DropListBox 185, 130, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_5_dropdown
	DropListBox 290, 130, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_5_dropdown
	DropListBox 55, 145, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_6_dropdown
	DropListBox 185, 145, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_6_dropdown
	DropListBox 290, 145, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_6_dropdown
	DropListBox 55, 160, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_7_dropdown
	DropListBox 185, 160, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_7_dropdown
	DropListBox 290, 160, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_7_dropdown
	DropListBox 55, 175, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_8_dropdown
	DropListBox 185, 175, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_8_dropdown
	DropListBox 290, 175, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_8_dropdown
	EditBox 75, 195, 55, 15, diet_date_last_exam
	DropListBox 135, 215, 35, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_treatment_plan_dropdown			'TODO: Handling for each scenario- each has it's own notification process/steps
	EditBox 270, 215, 55, 15, diet_length_diet
	DropListBox 55, 235, 60, 15, ""+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", diet_status_dropdown
	EditBox 50, 260, 290, 15, diet_comments
	PushButton 5, 280, 80, 15, "CM23.12- Special Diets", diet_link_CM_special_diet
    PushButton 95, 280, 115, 15, "Processing Special Diet Referrals", diet_SP_referrals
	Text 5, 5, 220, 10, diet_form_name
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, MAXIS_case_number
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 20, 40, 30, 10, "Member"
	Text 185, 40, 50, 10, "MFIP or MSA?"
	Text 235, 40, 85, 10, diet_mfip_msa_status
	Text 55, 60, 85, 10, "Select Applicable Diet"
	Text 185, 60, 95, 10, "Relationship between diets"
	Text 300, 60, 15, 10, "Ver"
	Text 30, 70, 20, 10, "Diet 1"
	Text 30, 85, 20, 10, "Diet 2"
	Text 30, 100, 20, 10, "Diet 3"
	Text 30, 115, 20, 10, "Diet 4"
	Text 30, 130, 20, 10, "Diet 5"
	Text 30, 145, 20, 10, "Diet 6"
	Text 30, 160, 20, 10, "Diet 7"
	Text 30, 175, 20, 10, "Diet 8"
	Text 15, 200, 60, 10, "Date of last exam"
	Text 15, 220, 115, 10, "Is person following treament plan?"
	Text 185, 220, 85, 10, "Length of Prescribed Diet"
	Text 15, 240, 40, 10, "Diet status?"
	Text 15, 265, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function
Dim diet_effective_date, diet_date_received, diet_member_number, diet_mfip_msa_status, diet_1_dropdown, diet_2_dropdown, diet_3_dropdown, diet_4_dropdown, diet_5_dropdown, diet_6_dropdown, diet_7_dropdown, diet_8_dropdown, diet_relationship_1_dropdown, diet_relationship_2_dropdown, diet_relationship_3_dropdown, diet_relationship_4_dropdown, diet_relationship_5_dropdown, diet_relationship_6_dropdown, diet_relationship_7_dropdown, diet_relationship_8_dropdown, diet_verif_1_dropdown, diet_verif_2_dropdown, diet_verif_3_dropdown, diet_verif_4_dropdown, diet_verif_5_dropdown, diet_verif_6_dropdown, diet_verif_7_dropdown, diet_verif_8_dropdown, diet_date_last_exam, diet_treatment_plan_dropdown, diet_status_dropdown, diet_length_diet, diet_comments	'Special Diet Variables

function get_footer_month_from_date(footer_month_variable, footer_year_variable, date_variable)

    footer_month_variable = DatePart("m", date_variable)
    footer_month_variable = Right("00" & footer_month_variable, 2)

    footer_year_variable = DatePart("yyyy", date_variable)
    footer_year_variable = Right(footer_year_variable, 2)
end function
'TODO- do i need to dim anything in this fuction?

function dialog_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs. 
	If form_count < Ubound(form_type_array, 2) and ButtonPressed = -1 Then ButtonPressed = next_btn	'If the enter button is selected the script will handle this as if Next was selected
	If form_count = Ubound(form_type_array, 2) and ButtonPressed = -1 Then ButtonPressed = complete_btn	'If the enter button is selected the script will handle this as if Complete was selected
	If ButtonPressed = next_btn Then form_count = form_count + 1	'If next is selected, it will iterate to the next form in the array and display this dialog
	If ButtonPressed = previous_btn Then form_count = form_count - 1	'If previous is selected, it will iterate to the previous form in the array and display this dialog
	If ButtonPressed >= 400 Then 'All forms are in the 400 range
		For i = 0 to Ubound(form_type_array, 2) 	'For/Next used to iterate through the array to display the correct dialog
			If ButtonPressed = asset_btn and form_type_array(form_type_const, i) = asset_form_name Then form_count = i 
			If ButtonPressed = atr_btn and form_type_array(form_type_const, i) = atr_form_name Then form_count = i 
			If ButtonPressed = arep_btn and form_type_array(form_type_const, i) = arep_form_name Then form_count = i 
			If ButtonPressed = change_btn and form_type_array(form_type_const, i) = change_form_name Then form_count = i 
			If ButtonPressed = evf_btn and form_type_array(form_type_const, i) = evf_form_name Then form_count = i 
			If ButtonPressed = hospice_btn and form_type_array(form_type_const, i) = hosp_form_name Then form_count = i 
			If ButtonPressed = iaa_btn and form_type_array(form_type_const, i) = iaa_form_name Then form_count = i 
			If ButtonPressed = iaa_ssi_btn and form_type_array(form_type_const, i) = iaa_ssi_form_name Then form_count = i 
			If ButtonPressed = ltc_1503_btn and form_type_array(form_type_const, i) = ltc_1503_form_name Then form_count = i 
			If ButtonPressed = mof_btn and form_type_array(form_type_const, i) = mof_form_name Then form_count = i 
			If ButtonPressed = mtaf_btn and form_type_array(form_type_const, i) = mtaf_form_name Then form_count = i 
			If ButtonPressed = psn_btn and form_type_array(form_type_const, i) = psn_form_name Then form_count = i 
			If ButtonPressed = sf_btn and form_type_array(form_type_const, i) = sf_form_name Then form_count = i 
			If ButtonPressed = diet_btn and form_type_array(form_type_const, i) = diet_form_name Then form_count = i 
		Next
	End If 
	
	If ButtonPressed = hosp_TE0207081_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.07.081%20HOSPICE%20CASES.pdf?csf=1&web=1&e=WgdqsC"
	If ButtonPressed = hosp_SP_hospice_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Hospice.aspx"
	If ButtonPressed = iaa_CM121203_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00121203"
	If ButtonPressed = iaa_te021214_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.12.14%20INTERIM%20ASSISTANCE%20REIMBURSEMENT%20INTERFACE.pdf?csf=1&web=1&e=tUXs96"
	If ButtonPressed = iaa_sp_pben_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/STAT_PBEN.aspx"
	If ButtonPressed = iaa_ssi_CM121203_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00121203"
	If ButtonPressed = iaa_ssi_te021214_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.12.14%20INTERIM%20ASSISTANCE%20REIMBURSEMENT%20INTERFACE.pdf?csf=1&web=1&e=tUXs96"
	If ButtonPressed = iaa_ssi_sp_pben_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/STAT_PBEN.aspx"
	If ButtonPressed = diet_link_CM_special_diet Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_002312"
	If ButtonPressed = diet_SP_referrals Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Processing_Special_Diet_Referral.aspx"
	If ButtonPressed = psn_TE1817_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2018.17%20ADULT%20GRH%20BASIS%20OF%20ELIGIBILITY.pdf?csf=1&web=1&e=7YWKmj"
	If ButtonPressed = psn_hss_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=DHS-316637"
	If ButtonPressed = psn_mhm_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg/Training_home_page.doc?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=dhs16_184936#em"
	If ButtonPressed = psn_hsss_btn	Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=dhs-289228"
end function 

function main_error_handling()	'Error handling for main dialog of forms
	If ButtonPressed = complete_btn Then 		'Error handling will display if the complete button is selected for the applicable dialogs.This happens at the end of the dialog series instead of at each dialog.
		For form_errors = 0 to Ubound(form_type_array, 2)
			If form_type_array(form_type_const, form_errors) = asset_form_name then 'Error handling for Asset Form 
				'TODO: Asset error reporting 
			End If

			If form_type_array(form_type_const, form_errors) = atr_form_name Then 'Error handling for ATR Form 
				If IsDate(atr_effective_date) = FALSE Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the Effective Date."
				If IsDate(atr_date_received) = FALSE Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				If atr_member_dropdown = "Select" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a member from the Member dropdown."
				If IsDate(atr_start_date) = FALSE Then  atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the Start Date."
				If IsDate(atr_end_date) = FALSE Then  atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the End Date."
				If atr_authorization_type = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a valid authorization type from the dropdown"
				If atr_contact_type = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a valid contact type from the dropdown"
				If atr_name = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter contact name"
				If trim(atr_address = "") OR trim(atr_city = "") OR trim(atr_state = "") OR trim(atr_zipcode = "") Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter the street address"
				If atr_phone_number = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter phone number"
				If (atr_eval_treat_checkbox = 0 and atr_coor_serv_checkbox = 0 and atr_elig_serv_checkbox = 0 and atr_court_checkbox = 0 and atr_other_checkbox = 0) Then atr_err_msg = atr_err_msg & vbNewLine & "* Must check at least one checkbox indicating use of requested record"
				If (atr_other_checkbox = checked and atr_other = "") Then err_msg = err_msg & vbNewLine & "* Other checkbox checked, specify details in the box below checkbox"
				If (atr_other <> "" and atr_other_checkbox = unchecked) Then atr_err_msg = atr_err_msg & vbNewLine & "* Other text field must be blank unless Other checkbox is checked"
			End If

			If form_type_array(form_type_const, form_errors) = arep_form_name then 'Error handling for AREP Form 
				If IsDate(arep_effective_date) = FALSE Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter a valid date for the Effective Date."
				If IsDate(arep_date_received) = FALSE Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				If trim(arep_name) = "" Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter the AREP's name."
				If arep_update_AREP_panel_checkbox = checked Then
					If trim(arep_street) = "" OR trim(arep_city) = "" OR trim(arep_zip) = "" Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter the street address of the AREP."
					If len(arep_name) > 37 Then arep_err_msg = arep_err_msg & vbNewLine & "* The AREP name is too long for MAXIS."
					If len(arep_street) > 44 Then arep_err_msg = arep_err_msg & vbNewLine & "* The AREP street is too long for MAXIS."
					If len(arep_city) > 15 Then arep_err_msg = arep_err_msg & vbNewLine & "* The AREP City is too long for MAXIS."
					If len(arep_state) > 2 Then arep_err_msg = arep_err_msg & vbNewLine & "* The AREP state is too long for MAXIS."
					If len(arep_zip) > 5 Then arep_err_msg = arep_err_msg & vbNewLine & "* The AREP zip is too long for MAXIS."
				End If
				If dhs_3437_checkbox = Checked Then arep_HC_AREP_checkbox = checked
				If HC_12729_checkbox = checked Then
					arep_SNAP_AREP_checkbox = checked
					arep_CASH_AREP_checkbox = checked
				End If
				If D405_checkbox = checked Then arep_SNAP_AREP_checkbox = checked
				If CAF_AREP_page_checkbox = checked Then
					arep_SNAP_AREP_checkbox = checked
					arep_CASH_AREP_checkbox = Checked
				End If
				If HCAPP_AREP_checkbox = checked Then arep_HC_AREP_checkbox = checked
				If power_of_attorney_checkbox = checked Then
					arep_SNAP_AREP_checkbox = checked
					arep_CASH_AREP_checkbox = Checked
					arep_HC_AREP_checkbox = checked
				End If
				'TODO: Look into arep received - -- there are currently 2 
				If IsDate(AREP_recvd_date) = False Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter the date the form was received."
				IF (arep_SNAP_AREP_checkbox <> checked AND arep_HC_AREP_checkbox <> checked AND arep_CASH_AREP_checkbox <> checked) THEN arep_err_msg = arep_err_msg & vbNewLine &"* Select a program"
				IF isdate(arep_signature_date) = false THEN arep_err_msg = arep_err_msg & vbNewLine & "* Enter a valid date for the date the form was signed/valid from."
				IF (arepTIKL_check = checked AND arep_signature_date = "") THEN arep_err_msg = arep_err_msg & vbNewLine & "* You have requested the script to TIKL based on the signature date but you did not enter the signature date."
				'TODO: Look into TIKLs
			End If

			If form_type_array(form_type_const, form_errors) = change_form_name  then 'Error handling for Change Form 
				If IsDate(trim(chng_effective_date)) = False OR Len(trim(chng_effective_date)) <> 10 Then chng_err_msg = chng_err_msg & vbNewLine & "* Date Effective field must be in the MM/DD/YYYY format."  ' Validate that Date Effective field is not empty and is in a proper date format
				If IsDate(trim(chng_date_received)) = False OR Len(trim(chng_date_received)) <> 10 Then chng_err_msg = chng_err_msg & vbNewLine & "* The Document Date field must be in the MM/DD/YYYY format."  ' Validate that Date Change Reported/Received field is not empty and is in a proper date format
				If trim(chng_address_notes) = "" AND trim(chng_household_notes) = "" AND trim(chng_asset_notes) = "" AND trim(chng_vehicles_notes) = "" AND trim(chng_income_notes) = "" AND trim(chng_shelter_notes) = "" AND trim(chng_other_change_notes) = "" THEN chng_err_msg = chng_err_msg & vbNewLine & "* All change reported fields are blank. At least one needs info."  ' Validate the Changes Reported fields to ensure that at least one field is filled in
				If trim(chng_actions_taken) = "" AND trim(chng_other_notes) = "" AND trim(chng_verifs_requested) = "" THEN chng_err_msg = chng_err_msg & vbNewLine & "* All of the Actions fields are blank. At least one need info."  ' Validate the Actions fields to ensure that at least one field is filled in
				If chng_notable_change = "" Then chng_err_msg = chng_err_msg & vbNewLine & "* Notable changes reported is blank, make a selection."
				If chng_changes_continue = "Select One:" THEN chng_err_msg = chng_err_msg & vbNewLine & "* Indicate whether changes will or will not continue next month."  ' Validate that worker selects option from dropdown list as to how long change will last
			
			End If

			If form_type_array(form_type_const, form_errors) = evf_form_name  then 'Error handling for EVF Form 
				IF IsDate(evf_effective_date) = FALSE THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter a valid Effective Date."
				IF IsDate(evf_date_received) = FALSE THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter a valid date for date the EVF was received."
				If EVF_status_dropdown = "Select one..." THEN evf_err_msg = evf_err_msg & vbCr & "* You must select the status of the EVF on the dropdown menu"		'checks that there is a date in the date received box
				IF evf_employer = "" THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter the employers name."  'checks if the employer name has been entered
				IF evf_client = "Select" THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter the MEMB information."  'checks if the client name has been entered
				IF evf_info = "Select one..." THEN evf_err_msg = evf_err_msg & vbCr & "* You must select if additional info was requested."  'checks if completed by employer was selected
				IF evf_info = "yes" and IsDate(evf_info_date) = FALSE THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter a valid date that additional info was requested."  'checks that there is a info request date entered if the it was requested
				IF evf_info = "yes" and evf_request_info = "" THEN evf_err_msg = evf_err_msg & vbCr & "* You must enter the method used to request additional info."		'checks that there is a method of inquiry entered if additional info was requested
				If evf_info = "no" and evf_request_info <> "" then evf_err_msg = evf_err_msg & vbCr & "* You cannot mark additional info as 'no' and have information requested."
				If evf_info = "no" and evf_info_date <> "" then evf_err_msg = evf_err_msg & vbCr & "* You cannot mark additional info as 'no' and have a date requested."
				'TODO: TIKL
				' If EVF_TIKL_checkbox = 1 and info <> "yes" then evf_err_msg = evf_err_msg & vbCr & "* Additional informaiton was not requested, uncheck the TIKL checkbox."
				' If ButtonPressed = 0 then evf_err_msg = "LOOP" & evf_err_msg
				' If skip_evf = TRUE Then
				' 	evf_form_received_checkbox = unchecked
				' 	evf_err_msg = ""
				' 	EVF_TIKL_checkbox = unchecked
				' End If
			End If
			If form_type_array(form_type_const, form_errors) = hosp_form_name  then 'Error handling for Hospice Form 
				If IsDate(hosp_effective_date) = FALSE Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter a valid date for the Effective Date." 
				If IsDate(hosp_date_received) = FALSE Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter a valid date for the Document Date." 
				If hosp_resident_name = "Select" Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Select the resident that is in hospice."
				If trim(hosp_name) = "" Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter the name of the Hospice the client entered."       'hospice name required
				If IsDate(hosp_entry_date) = FALSE Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter a valid date for the Hospice Entry."   'entry date also required
			End If

			If form_type_array(form_type_const, form_errors) = iaa_form_name then 'Error handling for IAA Form 
				IF IsDate(iaa_effective_date) = FALSE THEN iaa_err_msg = iaa_err_msg & vbCr & "* You must enter a valid Effective Date."
				IF IsDate(iaa_date_received) = FALSE THEN iaa_err_msg = iaa_err_msg & vbCr & "* You must enter a valid Document date."
				If iaa_member_dropdown = "Select" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Select the resident from the dropdown."
				If iaa_type_assistance = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Select type of interim assistance."
				If iaa_update_pben_checkbox = checked AND (iaa_benefit_type = "" or iaa_referral_date = "" or iaa_verification_dropdown = "" or iaa_date_applied_pben = "" or iaa_disposition_code_dropdown = "" or iaa_iaa_date = "") Then 
					iaa_err_msg = iaa_err_msg & vbNewLine & "* PBEN field requirements:"
					If iaa_benefit_type = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Select benefit type"
					If IsDate(iaa_referral_date) = FALSE Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter a valid referral date"
					If iaa_verification_dropdown = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Select verifiction type"
					If IsDate(iaa_date_applied_pben) = FALSE Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter a valid date for Date applied to PBEN"
					If iaa_disposition_code_dropdown = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Select disposition code"
					If iaa_date = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter valid IAA date"
				End If
			End If

			If form_type_array(form_type_const, form_errors) = iaa_ssi_form_name then 'Error handling for IAA-SSI Form
				IF IsDate(iaa_ssi_effective_date) = FALSE THEN iaa_err_msg = iaa_err_msg & vbCr & "* You must enter a valid Effective Date."
				IF IsDate(iaa_ssi_date_received) = FALSE THEN iaa_err_msg = iaa_err_msg & vbCr & "* You must enter a valid Document date."
				If iaa_ssi_member_dropdown = "Select" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "* Select the resident from the dropdown."
				If iaa_ssi_type_of_interim_assistance = "" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "* Select type of interim assistance."
				If iaa_ssi_update_pben_checkbox = checked AND (iaa_ssi_benefit_type = "" or iaa_ssi_referral_date = "" or iaa_ssi_verification_dropdown = "" or iaa_ssi_date_applied_pben = "" or iaa_ssi_disposition_code_dropdown = "" or iaa_ssi_iaa_date = "") Then 
					iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "* PBEN field requirements:"
					If iaa_ssi_benefit_type = "" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Select benefit type"
					If IsDate(iaa_ssi_referral_date) = FALSE Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Enter a valid referral date"
					If iaa_ssi_verification_dropdown = "" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Select verifiction type"
					If IsDate(iaa_ssi_date_applied_pben) = FALSE Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Enter a valid date for Date applied to PBEN"
					If iaa_ssi_disposition_code_dropdown = "" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Select disposition code"
					If iaa_ssi_iaa_date = "" Then iaa_ssi_err_msg = iaa_ssi_err_msg & vbNewLine & "  * Enter valid IAA date"
				End If
				
			End If
			If form_type_array(form_type_const, form_errors) = ltc_1503_form_name then 'Error handling for LTC 1503 Form
				'LTC 1503 -- didn't appear to be any error handling 
			End If
			If form_type_array(form_type_const, form_errors) = mof_form_name then 'Error handling for MOF Form 
			End If
			If form_type_array(form_type_const, form_errors) = mtaf_form_name then 'Error handling for MTAF Form
			End If
			If form_type_array(form_type_const, form_errors) = psn_form_name then 'Error handling for PSN Form
				IF IsDate(psn_date_received) = FALSE THEN psn_err_msg = psn_err_msg & vbCr & "* You must enter a valid Document Date."
				If psn_member_dropdown = "Select" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select the resident from the dropdown."
				If psn_section_1_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 1 make selection from dropdown."
				If psn_section_2_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 2 make selection from dropdown."
				If psn_section_3_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 3 make selection from dropdown."
				If psn_section_4_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 4 make selection from dropdown."
				If psn_section_5_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 5 make selection from dropdown."
				If psn_cert_prof = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Certified Professional or NA"
				If psn_facility = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Facilty name or NA"
				If psn_wreg_fs_pwe = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select FS PWE from dropdown"
				If psn_wreg_work_wreg_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select FSET Work Reg Status from dropdown"
				If psn_wreg_abawd_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select ABAWD Status from dropdown"
				If psn_wreg_ga_elig_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select GA Elig Basis Code from dropdown"
				If psn_disa_begin_date = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Disa Begin date"
				If psn_disa_end_date = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Disa Begin date"
				If psn_disa_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select Disa Status from dropdown"
				If psn_disa_verif = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select Verification from dropdown"
			End If

			If form_type_array(form_type_const, form_errors) = sf_form_name then 'Error handling for Shelter Form
			End If
			If form_type_array(form_type_const, form_errors) = diet_form_name then 'Error handling for Diet Form 
				If IsDate(diet_effective_date) = FALSE Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter a valid date for the Effective Date."
				If IsDate(diet_date_received) = FALSE Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				If diet_member_number = "Select" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select the resident for special diet."
	
				'Handling to ensure a relationship is selected if a diet has been entered on the line
				If diet_1_dropdown <>"" and diet_relationship_1_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 1 relationship" 
				If diet_2_dropdown <>"" and diet_relationship_2_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 2 relationship"
				If diet_3_dropdown <>"" and diet_relationship_3_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 3 relationship"
				If diet_4_dropdown <>"" and diet_relationship_4_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 4 relationship"
				If diet_5_dropdown <>"" and diet_relationship_5_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 5 relationship"
				If diet_6_dropdown <>"" and diet_relationship_6_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 6 relationship"
				If diet_7_dropdown <>"" and diet_relationship_7_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 7 relationship"
				If diet_8_dropdown <>"" and diet_relationship_8_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 8 relationship"

				'Handling to ensure a diet is selected if a relationship has been entered on the line
				If diet_relationship_1_dropdown <>"" and diet_1_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 1 diet"
				If diet_relationship_2_dropdown <>"" and diet_2_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 2 diet"
				If diet_relationship_3_dropdown <>"" and diet_3_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 3 diet"
				If diet_relationship_4_dropdown <>"" and diet_4_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 4 diet"
				If diet_relationship_5_dropdown <>"" and diet_5_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 5 diet"
				If diet_relationship_6_dropdown <>"" and diet_6_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 6 diet"
				If diet_relationship_7_dropdown <>"" and diet_7_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 7 diet"
				If diet_relationship_8_dropdown <>"" and diet_8_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet 8 diet"

				'Hnadling to ensure a verfication is selected if a diet and relationship have been entered on the same line
				If (diet_1_dropdown <> "" AND diet_relationship_1_dropdown <> "") AND diet_verif_1_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 1"
				If (diet_2_dropdown <> "" AND diet_relationship_2_dropdown <> "") AND diet_verif_2_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 2"
				If (diet_3_dropdown <> "" AND diet_relationship_3_dropdown <> "") AND diet_verif_3_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 3"
				If (diet_4_dropdown <> "" AND diet_relationship_4_dropdown <> "") AND diet_verif_4_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 4"
				If (diet_5_dropdown <> "" AND diet_relationship_5_dropdown <> "") AND diet_verif_5_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 5"
				If (diet_6_dropdown <> "" AND diet_relationship_6_dropdown <> "") AND diet_verif_6_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 6"
				If (diet_7_dropdown <> "" AND diet_relationship_7_dropdown <> "") AND diet_verif_7_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 7"
				If (diet_8_dropdown <> "" AND diet_relationship_8_dropdown <> "") AND diet_verif_8_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Verification for Diet 8"
				
					'Handling to limit diet selections to 1 type of protein.
				all_diet_string = "*"

				If diet_1_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_1_dropdown & "*"
				End If
				If diet_2_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_2_dropdown & "*"
				End If	
				If diet_3_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_3_dropdown & "*"
				End If 
				If diet_4_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_4_dropdown & "*"
				End If
				If diet_5_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_5_dropdown & "*"
				End If
				If diet_6_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_6_dropdown & "*"
				End If
				If diet_7_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_7_dropdown & "*"
				End If
				If diet_8_dropdown <>"" Then 
					all_diet_string = all_diet_string & diet_8_dropdown & "*"
				End If
				
				'TODO: Test on MFIP case handling to limit to 2 diets. Consider hiding extra boxes on dialog
				If (diet_mfip_msa_status = "MFIP-Active") OR (diet_mfip_msa_status = "MFIP-Pending") Then
					MsgBox "diet_mfip_msa_status" & diet_mfip_msa_status
						If diet_3_dropdown <>"" OR diet_4_dropdown <>"" OR diet_5_dropdown <>"" OR diet_6_dropdown <>"" OR diet_7_dropdown <>"" OR diet_8_dropdown <>"" Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have more than 2 diets for MFIP cases"
				End If

				If Instr(all_diet_string, "*01-High Protein*") AND Instr(all_diet_string, "*02-Controlled protein 40-60 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."
				If Instr(all_diet_string, "*01-High Protein*") AND Instr(all_diet_string,"*03-Controlled protein <40 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."
				If Instr(all_diet_string, "*02-Controlled protein 40-60 grams*") AND Instr(all_diet_string,"*03-Controlled protein <40 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."

				If IsDate(diet_date_last_exam) = FALSE Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter a valid date for Date of last exam."
				If diet_treatment_plan_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select dropdown indicating person is following treatment plan"
				If diet_length_diet = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter length of prescribed diet"
				If diet_status_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet Status"
				If diet_status_dropdown = "Denied" AND diet_comments = "" Then diet_err_msg = diet_err_msg & vbNewLine & "*Diet Denied, state reason & benefit end date in Comments"
				
			End If
		Next
		
		
	End If
	'Complete button triggers the error message to populate. Formatting error meessage to: Adds headers for each form if there are applicable errors
	If asset_err_msg <> "" Then err_msg = err_msg & vbNewLine & "ASSET DIALOG" & asset_err_msg & vbNewLine
	If atr_err_msg <> "" Then err_msg = err_msg & vbNewLine & "ATR DIALOG" & atr_err_msg & vbNewLine
	If arep_err_msg <> "" Then err_msg = err_msg & vbNewLine & "AREP DIALOG" & arep_err_msg & vbNewLine
	If chng_err_msg <> "" Then err_msg = err_msg & vbNewLine & "CHANGE DIALOG" & chng_err_msg & vbNewLine
	If evf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "EVF DIALOG" & evf_err_msg & vbNewLine
	If hosp_err_msg <> "" Then err_msg = err_msg & vbNewLine & "HOSPICE DIALOG" & hosp_err_msg & vbNewLine
	If iaa_err_msg <> "" Then err_msg = err_msg & vbNewLine & "IAA DIALOG" & iaa_err_msg & vbNewLine
	If iaa_ssi_err_msg <> "" Then err_msg = err_msg & vbNewLine & "IAA-SSI DIALOG" & iaa_ssi_err_msg & vbNewLine
	If ltc_1503_err_msg <> "" Then err_msg = err_msg & vbNewLine & "LTC 1503 DIALOG" & ltc_1503_err_msg & vbNewLine
	If mof_err_msg <> "" Then err_msg = err_msg & vbNewLine & "MOF DIALOG" & mof_err_msg & vbNewLine
	If mtaf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "MTAF DIALOG" & mtaf_err_msg & vbNewLine
	If psn_err_msg <> "" Then err_msg = err_msg & vbNewLine & "PSN DIALOG" & psn_err_msg & vbNewLine
	If sf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "SF DIALOG" & sf_err_msg & vbNewLine
	If diet_err_msg <> "" Then err_msg = err_msg & vbNewLine & "DIET DIALOG" & diet_err_msg & vbNewLine

	If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
end function

'Check for case number & footer & background
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
Call MAXIS_background_check

'DIALOG COLLECTING CASE, FOOTER MO/YR===========================================================================
Do
	DO
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 181, 90, "Case number dialog"
			EditBox 70, 5, 65, 15, MAXIS_case_number
			EditBox 70, 25, 30, 15, MAXIS_footer_month
			EditBox 105, 25, 30, 15, MAXIS_footer_year
			EditBox 70, 45, 100, 15, worker_signature
			ButtonGroup ButtonPressed
				OkButton 65, 70, 50, 15
				CancelButton 120, 70, 50, 15
			Text 20, 10, 50, 10, "Case number: "
			Text 20, 30, 45, 10, "Footer month:"
			Text 5, 50, 60, 10, "Worker signature:"
		EndDialog


		dialog Dialog1	'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF IsNumeric(MAXIS_footer_month) = FALSE OR IsNumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbNewLine &  "* You must type a valid footer month and year."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


Call Generate_Client_List(HH_Memb_DropDown, "Select")         'filling the dropdown with ALL of the household members


'DIALOGS COLLECTING FORM SELECTION===========================================================================
'TODO: Handle for duplicate selection
Do							'Do Loop to cycle through dialog as many times as needed until all desired forms are added
	Do
		Do
			err_msg = ""
			Dialog1 = "" 			'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
				DropListBox 30, 30, 180, 15, ""+chr(9)+asset_form_name+chr(9)+atr_form_name+chr(9)+arep_form_name+chr(9)+change_form_name+chr(9)+evf_form_name+chr(9)+hosp_form_name+chr(9)+iaa_form_name+chr(9)+iaa_ssi_form_name+chr(9)+ltc_1503_form_name+chr(9)+mof_form_name+chr(9)+mtaf_form_name+chr(9)+psn_form_name+chr(9)+sf_form_name+chr(9)+diet_form_name, Form_type
				ButtonGroup ButtonPressed
				PushButton 225, 30, 35, 10, "Add", add_button
				PushButton 225, 60, 35, 10, "All Forms", all_forms
				PushButton 155, 215, 40, 15, "Clear", clear_button
				OkButton 205, 215, 40, 15
				CancelButton 255, 215, 40, 15
				GroupBox 5, 5, 280, 70, "Directions: For each document received either:"
				Text 15, 15, 275, 10, "1. Select document from dropdown, then select Add button. Repeat for each form."
				Text 10, 45, 15, 10, "OR"
				Text 15, 60, 180, 10, "2. Select All Forms to select forms via checkboxes."
				GroupBox 45, 85, 210, 125, "Documents Selected"
				y_pos = 95			'defining y_pos so that we can dynamically add forms to the dialog as they are selected
				
				For form = 0 to UBound(form_type_array, 2) 'Writing form name by incrementing to the next value in the array. For/next must be within dialog so it knows where to write the information. 
					'MsgBox form_type_array(form_type_const, form) 'TEST
					'MsgBox form_type_array(btn_name_const, form) 'TEST
					'MsgBox form_type_array(btn_number_const, form) 'TEST
					'MsgBox "Ubound" & UBound(form_type_array, 2) 'TEST
						Text 55, y_pos, 195, 10, form_type_array(form_type_const, form)
						y_pos = y_pos + 10					'Increasing y_pos by 10 before the next form is written on the dialog
				Next
				EndDialog								'Dialog handling	
				dialog Dialog1 							'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation
				
			'This limits the quantity of each form to 1. Only adds the form name to the array if it's not already in there. If it's already in the array, it does not add it to the array. 
			If ButtonPressed = add_button Then 	'Add button kicks off this evaluation 
				If form_type <> "" Then 		'Must have a form selected
				Form_string = form_type 		'Setting the form name equal to a string
					'MsgBox "all_form_array" & all_form_array
					If instr(all_form_array, "*" & form_string & "*") Then	
						add_to_array = false	'If the string is found in the array, it won't add the form to the array
						'MsgBox "add_to_array = false"

					Else 
						add_to_array = true 	'If the string is not found in the array, it will add the form to the array
						'MsgBox "add_to_array = true"

					End If
				End If
			
				If add_to_array = True Then			'Defining the steps to take if the form should be added to the array
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = Form_type		'Storing form name in the array		
					form_count = form_count + 1 	
					all_form_array = all_form_array & form_string & "*" 'Adding form name to form name string
				' ElseIF add_to_array = False then 
				' 	false_count = false_count + 1
				End If 
			End If

			
			'MsgBox "all form array string" & all_form_array '= split(all_form_array, "*")

			'This work for handling the adding of each form - this allows you to add more than one of each form 
			' If ButtonPressed = add_button and form_type <> "" Then				'If statement to know when to store the information in the array
			' 	ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
			' 	form_type_array(form_type_const, form_count) = Form_type		'Storing form name in the array		
			' 	form_count= form_count + 1 										'incrementing in the array
			' End If
				
			If ButtonPressed = clear_button Then 'Clear button wipes out any selections already made so the user can reselect correct forms.
				ReDim form_type_array(the_last_const, form_count)		
				form_count = 0							'Reset the form count to 0 so that y_pos resets to 95. 
				Form_string = ""						'Reset string to nothing
				add_to_array = ""						'reset to nothing
				all_form_array = "*"					'Reset string to *		
				 
				'Future Iteration: Counting quantity of each form
				' asset_checkbox = unchecked				'Resetting checkboxes to unchecked
				' atr_checkbox = unchecked				'Resetting checkboxes to unchecked
				' arep_checkbox = unchecked				'Resetting checkboxes to unchecked
				' change_checkbox = unchecked				'Resetting checkboxes to unchecked
				' evf_checkbox = unchecked				'Resetting checkboxes to unchecked
				' hospice_checkbox = unchecked			'Resetting checkboxes to unchecked
				' iaa_checkbox = unchecked				'Resetting checkboxes to unchecked
				' iaa_ssi_checkbox = unchecked			'Resetting checkboxes to unchecked
				' ltc_1503_checkbox = unchecked			'Resetting checkboxes to unchecked
				' mof_checkbox = unchecked				'Resetting checkboxes to unchecked
				' mtaf_checkbox = unchecked				'Resetting checkboxes to unchecked
				' psn_checkbox = unchecked				'Resetting checkboxes to unchecked
				' shelter_checkbox = unchecked			'Resetting checkboxes to unchecked
				' diet_checkbox = unchecked				'Resetting checkboxes to unchecked
				' form_type = ""							'Resetting dropdown to blank


				'MsgBox "form string" & form_string
				'MsgBox "all form array" & all_form_array
				'
				'Future Iteration: Counting quantity of each form
				' asset_count 	= 0 
				' atr_count 		= 0 
				' arep_count 		= 0 
				' change_count 	= 0
				' evf_count		= 0 
				' hosp_count		= 0 
				' iaa_count		= 0 
				' iaa_ssi_count	= 0
				' ltc_1503_count	= 0
				' mof_count		= 0 
				' mtaf_count		= 0 
				' psn_count		= 0 
				' sf_count		= 0 
				' diet_count		= 0
				
				'MsgBox "all form array" & all_form_array
			'	MsgBox "form type" & form_type 'TEST
				MsgBox "Form selections cleared." 'Notify end user that entries were cleared.
				'MsgBox "all_form_array" & all_form_array
			End If

			If ButtonPressed = add_button Then 'Handles for duplicates and no forms selected from dropdown.
				If form_type <> "" Then 
					If add_to_array = FALSE Then err_msg = err_msg & vbNewLine & "Form already added, make a different form selection."
				End If
				If form_type = "" Then err_msg = err_msg & vbNewLine & "No form selected, make form selection."
			End If
			If form_count = 0 and ButtonPressed = Ok Then err_msg = "-Add forms to process or select cancel to exit script"		'If form_count = 0, then no forms have been added to doc rec to be processed.	
			If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	If ButtonPressed = all_forms Then		'Opens Dialog with checkbox selection for each form
		Do
			Do
				ReDim form_type_array(the_last_const, form_count)		'Resetting any selections already made so the user can reselect correct forms using different format.
				'form_type_array(form_type_const, form_count) = Form_type
                form_count = 0							'Resetting the form count to 0 so that y_pos resets to 95. 
				Form_string = ""						'Resetting string to nothing 
				all_form_array = "*"						'Resetting list of strings to * 
				add_to_array = ""

				'Future Iteration - carries values selected from drop down through to checkbox feature
				' If instr(all_form_array, asset_form_name) Then asset_checkbox = checked 
				' If instr(all_form_array, atr_form_name) Then atr_checkbox = checked 
				' If instr(all_form_array, arep_form_name) Then arep_checkbox = checked 
				' If instr(all_form_array, change_form_name) Then change_checkbox = checked 
				' If instr(all_form_array, evf_form_name) Then evf_checkbox = checked 
				' If instr(all_form_array, hosp_form_name) Then hospice_checkbox = checked 
				' If instr(all_form_array, iaa_form_name) Then iaa_checkbox = checked 
				' If instr(all_form_array, iaa_ssi_form_name) Then iaa_ssi_checkbox = checked 
				' If instr(all_form_array, ltc_1503_form_name) Then ltc_1503_checkbox = checked 
				' If instr(all_form_array, mof_form_name) Then mof_checkbox = checked 
				' If instr(all_form_array, mtaf_form_name) Then mtaf_checkbox = checked 
				' If instr(all_form_array, psn_form_name) Then psn_checkbox = checked 
				' If instr(all_form_array, sf_form_name) Then shelter_checkbox = checked 
				' If instr(all_form_array, diet_form_name) Then diet_checkbox = checked 


				err_msg = ""
				Dialog1 = "" 'Blanking out previous dialog detail
				BeginDialog Dialog1, 0, 0, 196, 200, "Document Selection"
					CheckBox 15, 20, 160, 10, asset_form_name, asset_checkbox
					CheckBox 15, 30, 160, 10, atr_form_name, atr_checkbox
					CheckBox 15, 40, 160, 10, arep_form_name, arep_checkbox
					CheckBox 15, 50, 160, 10, change_form_name, change_checkbox
					CheckBox 15, 60, 160, 10, evf_form_name, evf_checkbox
					CheckBox 15, 70, 160, 10, hosp_form_name, hospice_checkbox
					CheckBox 15, 80, 160, 10, iaa_form_name, iaa_checkbox
					CheckBox 15, 90, 160, 10, iaa_ssi_form_name, iaa_ssi_checkbox
					CheckBox 15, 100, 160, 10, ltc_1503_form_name, ltc_1503_checkbox
					CheckBox 15, 110, 160, 10, mof_form_name, mof_checkbox
					CheckBox 15, 120, 160, 10, mtaf_form_name, mtaf_checkbox
					CheckBox 15, 130, 160, 10, psn_form_name, psn_checkbox
					CheckBox 15, 140, 170, 10, sf_form_name, shelter_checkbox
					CheckBox 15, 150, 175, 10, diet_form_name, diet_checkbox
					ButtonGroup ButtonPressed
					OkButton 95, 180, 45, 15
					CancelButton 150, 180, 40, 15
					Text 5, 5, 200, 10, "Select documents received, then Ok."
				EndDialog
				dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation

				
				
				'Capturing form name in array based on checkboxes selected 
				If asset_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = asset_form_name 
					form_count= form_count + 1 
				End If

				If atr_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = atr_form_name
					form_count= form_count + 1 
				End If

				If arep_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = arep_form_name
					form_count= form_count + 1 
				End If

				If change_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = change_form_name
					form_count= form_count + 1 
				End If
				If evf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = evf_form_name
					form_count= form_count + 1 
				End If
				If hospice_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = hosp_form_name
					form_count= form_count + 1 
				End If
				If iaa_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = iaa_form_name
					form_count= form_count + 1 
				End If
				If iaa_ssi_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = iaa_ssi_form_name
					form_count= form_count + 1 
				End If
				If ltc_1503_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = ltc_1503_form_name
					form_count= form_count + 1 
				End If
				If mof_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = mof_form_name
					form_count= form_count + 1 
				End If
				If mtaf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = mtaf_form_name
					form_count= form_count + 1 
				End If
				If psn_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = psn_form_name
					form_count= form_count + 1 
				End If
				If shelter_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = sf_form_name
					form_count= form_count + 1 
				End If
				If diet_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = diet_form_name
					form_count= form_count + 1 
				End If
			
				'MsgBox "all form array string" & all_form_array 
					
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and iaa_ssi_checkbox = unchecked and ltc_1503_checkbox = unchecked and mof_checkbox = unchecked and mtaf_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""	
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If		
Loop Until ButtonPressed = Ok

'TODO: Add in any additonal readscreens etc.
'MAXIS NAVIGATION ===========================================================================
Call MAXIS_background_check
For maxis_panel_read = 0 to Ubound(form_type_array, 2)
	If form_type_array(form_type_const, maxis_panel_read) = arep_form_name Then 'MAXIS NAVIGATION FOR AREP
		Do
			Call navigate_to_MAXIS_screen("STAT", "AREP")
			EMReadScreen nav_check, 4, 2, 53
			EMWaitReady 0, 0
		Loop until nav_check = "AREP"

		arep_update_AREP_panel_checkbox = checked
		AREP_recvd_date = arep_date_received

		EMReadScreen arep_name, 37, 4, 32
		arep_name = replace(arep_name, "_", "")
		If arep_name <> "" Then
			EMReadScreen arep_street_one, 22, 5, 32
			EMReadScreen arep_street_two, 22, 6, 32
			EMReadScreen arep_city, 15, 7, 32
			EMReadScreen arep_state, 2, 7, 55
			EMReadScreen arep_zip, 5, 7, 64

			arep_street_one = replace(arep_street_one, "_", "")
			arep_street_two = replace(arep_street_two, "_", "")
			arep_street = arep_street_one & " " & arep_street_two
			arep_street = trim(arep_street)
			arep_city = replace(arep_city, "_", "")
			arep_state = replace(arep_state, "_", "")
			arep_zip = replace(arep_zip, "_", "")

			EMReadScreen arep_phone_one, 14, 8, 34
			EMReadScreen arep_ext_one, 3, 8, 55
			EMReadScreen arep_phone_two, 14, 9, 34
			EMReadScreen arep_ext_two, 3, 8, 55

			arep_phone_one = replace(arep_phone_one, ")", "")
			arep_phone_one = replace(arep_phone_one, "  ", "-")
			arep_phone_one = replace(arep_phone_one, " ", "-")
			If arep_phone_one = "___-___-____" Then arep_phone_one = ""

			arep_phone_two = replace(arep_phone_two, ")", "")
			arep_phone_two = replace(arep_phone_two, "  ", "-")
			arep_phone_two = replace(arep_phone_two, " ", "-")
			If arep_phone_two = "___-___-____" Then arep_phone_two = ""

			arep_ext_one = replace(arep_ext_one, "_", "")
			arep_ext_two = replace(arep_ext_two, "_", "")

			EMReadScreen arep_forms_to_arep, 1, 10, 45
			EMReadScreen arep_mmis_mail_to_arep, 1, 10, 77

			If arep_forms_to_arep = "Y" Then arep_forms_to_arep_checkbox = checked
			If arep_mmis_mail_to_arep = "Y" Then arep_mmis_mail_to_arep_checkbox = checked

			arep_update_AREP_panel_checkbox = unchecked
			MsgBox "pause after read screen"
		End If
	End If

	If form_type_array(form_type_const, maxis_panel_read) = hosp_form_name Then	'MAXIS NAVIGATION FOR HOSPICE READ SCREEN
		Call navigate_to_MAXIS_screen("CASE", "NOTE")
		note_row = 5                                'beginning of listed case notes
		one_year_ago = DateAdd("yyyy", -1, date)    'we will look back 1 year
		Do
			EMReadScreen note_date, 8, note_row, 6      'reading the date
			EMReadScreen note_title, 55, note_row, 25   'reading the header
			note_title = trim(note_title)

			If left(note_title, 41) = "*** HOSPICE TRANSACTION FORM RECEIVED ***" Then      'if the note is for a Hospice form
				Call write_value_and_transmit("X", note_row, 3)	'open the note

				this_row = 5            'this MAXIS is the top of the note body
				Do
					EMReadScreen note_line, 78, this_row, 3     'reading each line
					note_line = trim(note_line)                 'Each of the lines will have the header look at to see if we can autofill information

					If  left(note_line, 9) = "* Client:" Then
						hosp_resident_name = right(note_line, len(note_line) - 9)
						hosp_resident_name = trim(hosp_resident_name)

					ElseIf left(note_line, 15) = "* Hospice Name:" Then
						hosp_name = right(note_line, len(note_line) - 15)
						hosp_name = trim(hosp_name)

					ElseIf left(note_line, 13) = "* NPI Number:" Then
						hosp_npi_number = right(note_line, len(note_line) - 13)
						hosp_npi_number = trim(hosp_npi_number)

					ElseIf left(note_line, 16) = "* Date of Entry:" Then
						hosp_entry_date = right(note_line, len(note_line) - 16)
						hosp_entry_date = trim(hosp_entry_date)

					ElseIf left(note_line, 12) = "* Exit Date:" Then
						hosp_exit_date = right(note_line, len(note_line) - 12)
						hosp_exit_date = trim(hosp_exit_date)

					ElseIf left(note_line, 26) = "* MMIS not updated due to:" Then
						hosp_reason_not_updated = right(note_line, len(note_line) - 26)
						hosp_reason_not_updated = trim(hosp_reason_not_updated)
					End If
					If this_row = 18 Then       'this is the bottom of the note, will go to the next page if possible
						PF8
						EMReadScreen check_for_end, 9, 24, 14   'if we try to PF8 and it doesn't go down, a message happens at the bottom
						If check_for_end = "LAST PAGE" Then
							PF3             'leaving the note
							Exit Do         'don't need to look at any more of the note
						End If
						this_row = 4        'if the message isn't there reset the row to the top
					End If
					this_row = this_row + 1     'go to the next row
					If note_line = "" Then PF3  'if it is blank - the note is over and we need to leave the note
				Loop until note_line = ""

				Exit Do     'if a HOSPICE note is found, we don't need to look at more notes
			End If
			IF note_date = "        " then Exit Do      'if the end of the list is reached we leave the loop
			note_row = note_row + 1
			IF note_row = 19 THEN       'going to the next page of notes
				PF8
				note_row = 5
			END IF
			EMReadScreen next_note_date, 8, note_row, 6
			IF next_note_date = "        " then Exit Do
		Loop until datevalue(next_note_date) < one_year_ago 'looking ahead at the next case note kicking out the dates before app'

		If hospice_exit_date <> "" Then     'if there is an exit date in the note found then we don't want to use the information from that note
			hosp_resident_name = ""          'since if they exited already - the HOSPICE will be different - resetting these variables to NOT fill
			hosp_name = ""
			hosp_npi_number = ""
			hosp_entry_date = ""
			hosp_exit_date = ""
			hosp_reason_not_updated = ""
		End If
		Do
			Call navigate_to_MAXIS_screen ("STAT", "MEMB")      'Going to MEMB for M01 to see if there is a date of death - as that would be the exit date
			EMReadScreen nav_check, 4, 2, 48
			EMWaitReady 0, 0
		Loop until nav_check = "MEMB"
		EMReadScreen date_of_death, 10, 19, 42
		date_of_death = replace(date_of_death, " ", "/")
		If IsDate(date_of_death) = TRUE Then hosp_exit_date = date_of_death
	End If

	If form_type_array(form_type_const, maxis_panel_read) = diet_form_name Then	'MAXIS NAVIGATION FOR DIET- Reading status of programs
		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
		If Instr(list_active_programs, "MSA") Then 
			diet_mfip_msa_status = "MSA-Active"
		ElseIf Instr(list_active_programs, "MFIP") Then 
			diet_mfip_msa_status = "MFIP-Active"
		ElseIf Instr(list_pending_programs, "MSA") Then 
			diet_mfip_msa_status = "MSA-Pending"
		ElseIf Instr(list_pending_programs, "MFIP") Then 
			diet_mfip_msa_status = "MFIP-Pending"
		Else 
			diet_mfip_msa_status = "Not Active/Pending"
		End If
		'MsgBox "dialog mfip/msa" & diet_mfip_msa_status
	End IF 

	If form_type_array(form_type_const, maxis_panel_read) = psn_form_name Then	'MAXIS NAVIGATION FOR PSN- reading necessary panels

		'Read WREG Panel
		Do
			Call navigate_to_MAXIS_screen("STAT", "WREG")
			EMReadScreen nav_check, 4, 2, 48
			EMWaitReady 0, 0
		Loop until nav_check = "WREG"
		Call write_value_and_transmit(Left(psn_member_dropdown, 2), 20, 76)
		EMReadScreen psn_wreg_fs_pwe, 1, 6, 68						'Read fs pwe
		EMReadScreen psn_wreg_work_wreg_status, 2, 8, 50			'Read work wreg status
		EMReadScreen psn_wreg_abawd_status, 2, 13, 50				'Read abawd status
		EMReadScreen psn_wreg_ga_elig_status, 2, 15, 50				'Read ga eligibility status
		psn_wreg_fs_pwe = replace(psn_wreg_fs_pwe, "_", "")
		psn_wreg_work_wreg_status = replace(psn_wreg_work_wreg_status, "__", "")
		psn_wreg_abawd_status = replace(psn_wreg_abawd_status, "__", "")
		psn_wreg_ga_elig_status = replace(psn_wreg_ga_elig_status, "__", "")
		'MsgBox "psn_wreg_fs_pwe" & psn_wreg_fs_pwe
		'MsgBox "psn_wreg_work_wreg_status" & psn_wreg_work_wreg_status
		'MsgBox	"psn_wreg_abawd_status" & psn_wreg_abawd_status
		'MsgBox	"psn_wreg_ga_elig_status" & psn_wreg_ga_elig_status

		'List of Work Wreg Status
		If psn_wreg_work_wreg_status = "03" Then psn_wreg_work_wreg_status = "03-Unfit for Employment"
		If psn_wreg_work_wreg_status = "04" Then psn_wreg_work_wreg_status = "04-Resp for Care of Incapacitated Person"
		If psn_wreg_work_wreg_status = "05" Then psn_wreg_work_wreg_status = "05-Age 60 or Older"
		If psn_wreg_work_wreg_status = "06" Then psn_wreg_work_wreg_status = "06-Under Age 16"
		If psn_wreg_work_wreg_status = "07" Then psn_wreg_work_wreg_status = "07-Age 16-17, Living w/ Caregiver"
		If psn_wreg_work_wreg_status = "08" Then psn_wreg_work_wreg_status = "08-Resp for Care of Child under 6"
		If psn_wreg_work_wreg_status = "09" Then psn_wreg_work_wreg_status = "09-Empl 30 hrs/wk or Earnings of 30 hrs/wk"
		If psn_wreg_work_wreg_status = "10" Then psn_wreg_work_wreg_status = "10-Matching Grant Participant"
		If psn_wreg_work_wreg_status = "11" Then psn_wreg_work_wreg_status = "11-Receiving or Applied for UI"
		If psn_wreg_work_wreg_status = "12" Then psn_wreg_work_wreg_status = "12-Enrolled in School, Training, or Higher Ed"
		If psn_wreg_work_wreg_status = "13" Then psn_wreg_work_wreg_status = "13-Participating in CD Program"
		If psn_wreg_work_wreg_status = "14" Then psn_wreg_work_wreg_status = "14-Receiving MFIP"
		If psn_wreg_work_wreg_status = "20" Then psn_wreg_work_wreg_status = "20-Pending/Receiving DWP"
		If psn_wreg_work_wreg_status = "15" Then psn_wreg_work_wreg_status = "15-Age 16-17, NOT Living w/ Caregiver"
		If psn_wreg_work_wreg_status = "16" Then psn_wreg_work_wreg_status = "16-50-59 Years Old"
		If psn_wreg_work_wreg_status = "17" Then psn_wreg_work_wreg_status = "17-Receiving RCA or GA"
		If psn_wreg_work_wreg_status = "21" Then psn_wreg_work_wreg_status = "21-Resp for Care of Child under 18"
		If psn_wreg_work_wreg_status = "23" Then psn_wreg_work_wreg_status = "23-Pegnant"

		'List of ABAWD Status 
		If psn_wreg_abawd_status = "01" Then psn_wreg_abawd_status = "01-Work Reg Exempt"
		If psn_wreg_abawd_status = "02" Then psn_wreg_abawd_status = "02-Under Age 18"
		If psn_wreg_abawd_status = "03" Then psn_wreg_abawd_status = "03-Age 50 or Over"
		If psn_wreg_abawd_status = "04" Then psn_wreg_abawd_status = "04-Caregiver of Minor Child"
		If psn_wreg_abawd_status = "05" Then psn_wreg_abawd_status = "05-Pregnant"
		If psn_wreg_abawd_status = "06" Then psn_wreg_abawd_status = "06-Employed Avg of 20 hrs/wk"
		If psn_wreg_abawd_status = "07" Then psn_wreg_abawd_status = "07-Work Experience Participant"
		If psn_wreg_abawd_status = "08" Then psn_wreg_abawd_status = "08-ther E&T Services"
		If psn_wreg_abawd_status = "09" Then psn_wreg_abawd_status = "09-Resides in a Waivered Area"
		If psn_wreg_abawd_status = "10" Then psn_wreg_abawd_status = "10-ABAWD Counted Month"
		If psn_wreg_abawd_status = "11" Then psn_wreg_abawd_status = "11-2nd-3rd Month Period of Elig"
		If psn_wreg_abawd_status = "12" Then psn_wreg_abawd_status = "12-RCA or GA Recipient"
		If psn_wreg_abawd_status = "13" Then psn_wreg_abawd_status = "13-ABAWD Banked Months"

		'List of Elig Status
		If psn_wreg_ga_elig_status = "04" Then psn_wreg_ga_elig_status = "04-Permanent Ill or Incap"
		If psn_wreg_ga_elig_status = "05" Then psn_wreg_ga_elig_status = "05-Temporary Ill or Incap"
		If psn_wreg_ga_elig_status = "06" Then psn_wreg_ga_elig_status = "06-Care of Ill or Incap Mbr"
		If psn_wreg_ga_elig_status = "07" Then psn_wreg_ga_elig_status = "07-Requires Services In Residence"
		If psn_wreg_ga_elig_status = "09" Then psn_wreg_ga_elig_status = "09-Mntl Ill or Dev Disabled"
		If psn_wreg_ga_elig_status = "10" Then psn_wreg_ga_elig_status = "10-SSI/RSDI Pend"
		If psn_wreg_ga_elig_status = "11" Then psn_wreg_ga_elig_status = "11-Appealing SSI/RSDI Denial"
		If psn_wreg_ga_elig_status = "12" Then psn_wreg_ga_elig_status = "12-Advanced Age"
		If psn_wreg_ga_elig_status = "13" Then psn_wreg_ga_elig_status = "13-Learning Disability"
		If psn_wreg_ga_elig_status = "17" Then psn_wreg_ga_elig_status = "17-Protect/Court Ordered"
		If psn_wreg_ga_elig_status = "20" Then psn_wreg_ga_elig_status = "20-Age 16 or 17 SS Approval"
		If psn_wreg_ga_elig_status = "25" Then psn_wreg_ga_elig_status = "25-Emancipated Minor"
		If psn_wreg_ga_elig_status = "28" Then psn_wreg_ga_elig_status = "28-Unemployable"
		If psn_wreg_ga_elig_status = "29" Then psn_wreg_ga_elig_status = "29-Displaced Hmkr (Ft Student)"
		If psn_wreg_ga_elig_status = "30" Then psn_wreg_ga_elig_status = "30-Minor w/ Adult Unrelated"
		If psn_wreg_ga_elig_status = "32" Then psn_wreg_ga_elig_status = "32-ESL, Adult/HS At least half time"
		If psn_wreg_ga_elig_status = "35" Then psn_wreg_ga_elig_status = "35-Drug/Alcohol Addiction (DAA)"
		If psn_wreg_ga_elig_status = "99" Then psn_wreg_ga_elig_status = "99-No Elig Basis"

		
		'Read Disa Panel
		Do
			Call navigate_to_MAXIS_screen("STAT", "DISA")
			EMReadScreen nav_check, 4, 2, 45
			EMWaitReady 0, 0
		Loop until nav_check = "DISA"
  		Call write_value_and_transmit(Left(psn_member_dropdown, 2), 20, 76)
		EMReadScreen psn_disa_begin_date, 10, 6, 47
		EMReadScreen psn_disa_end_date, 10, 6, 69
		EMReadScreen psn_disa_status, 2 , 11, 59
		EMReadScreen psn_disa_verif, 1, 11, 69
		psn_disa_begin_date = replace(psn_disa_begin_date,  "__ __ ____", "")
		psn_disa_end_date = replace(psn_disa_end_date,  "__ __ ____", "")
		psn_disa_begin_date = replace(psn_disa_begin_date, " ", "/")
		psn_disa_end_date = replace(psn_disa_end_date, " ", "/")
		psn_disa_status = replace(psn_disa_status, "__", "")
		psn_disa_verif = replace(psn_disa_verif, "_", "")
		'MsgBox "psn_disa_begin_date" & psn_disa_begin_date & "psn_disa_end_date" & psn_disa_end_date & "psn_disa_status" & psn_disa_status & "psn_disa_verif" & psn_disa_verif

		'List of Disa Status
		If psn_disa_status = "01" Then psn_disa_status = "01-RSDI Only Disability"
		If psn_disa_status = "02" Then psn_disa_status = "02-RSDI Only Blindness"
		If psn_disa_status = "03" Then psn_disa_status = "03-SSI, SSI/RSDI Disability"
		If psn_disa_status = "04" Then psn_disa_status = "04-SSI, SSI/RSDI Blindness"
		If psn_disa_status = "06" Then psn_disa_status = "06-SMRT/SSA Pend"
		If psn_disa_status = "08" Then psn_disa_status = "08-SMRT Certified Blindness"
		If psn_disa_status = "09" Then psn_disa_status = "09-Ill/Incapacity"
		If psn_disa_status = "10" Then psn_disa_status = "10-SMRT Certified Disability"

		'List of Disa Verification
		If psn_disa_verif = "1" Then psn_disa_verif = "1-DHS161/Dr Stmt"
		If psn_disa_verif = "2" Then psn_disa_verif = "2-SMRT Certified"
		If psn_disa_verif = "3" Then psn_disa_verif = "3-Certified For RSDI or SSI"
		If psn_disa_verif = "6" Then psn_disa_verif = "6-Other Document"
		If psn_disa_verif = "7" Then psn_disa_verif = "7-Professional Stmt of Need"
		If psn_disa_verif = "N" Then psn_disa_verif = "N-No Ver Prvd"

	'TODO: Only need to read vendor from FACI Panel and plug this into the dialog for the Facility  
		Do
			Call navigate_to_MAXIS_screen("STAT", "FACI")		'Navigate to FACI 
			EMReadScreen nav_check, 4, 2, 44
			EMWaitReady 0, 0
		Loop until nav_check = "FACI"
		EMWriteScreen Left(psn_member_dropdown, 2), 20, 76
		Call write_value_and_transmit("01", 20, 79)		'Make sure we are on page 01 for the member to start
			EMReadScreen existance_check, 1, 2, 73
			stat_faci_exists = True		'TODO: DO we need this?
			If existance_check = "0" Then stat_faci_exists = False
			'MsgBox "before if statement"
			If stat_faci_exists = True Then
				'MsgBox "stat_faci_exists = True"
				Do
					EMReadScreen FACI_current_panel, 1, 2, 73
					EMReadScreen FACI_total_check, 1, 2, 78
					EMReadScreen in_year_check_01, 4, 14, 53
					EMReadScreen in_year_check_02, 4, 15, 53
					EMReadScreen in_year_check_03, 4, 16, 53
					EMReadScreen in_year_check_04, 4, 17, 53
					EMReadScreen in_year_check_05, 4, 18, 53
					EMReadScreen out_year_check_01, 4, 14, 77
					EMReadScreen out_year_check_02, 4, 15, 77
					EMReadScreen out_year_check_03, 4, 16, 77
					EMReadScreen out_year_check_04, 4, 17, 77
					EMReadScreen out_year_check_05, 4, 18, 77
					'MsgBox "FACI_current_panel" & FACI_current_panel & "FACI_total_check" & FACI_total_check
		
					If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or _
					(in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
						currently_in_FACI = True
						'MsgBox "in_year_check_01" & in_year_check_01 & "out_year_check_01" & out_year_check_01
						'MsgBox "currently_in_FACI = True"
						If in_year_check_01 <> "____" and out_year_check_01 = "____" Then faci_row = 14
						If in_year_check_02 <> "____" and out_year_check_02 = "____" Then faci_row = 15
						If in_year_check_03 <> "____" and out_year_check_03 = "____" Then faci_row = 16
						If in_year_check_04 <> "____" and out_year_check_04 = "____" Then faci_row = 17
						If in_year_check_05 <> "____" and out_year_check_05 = "____" Then faci_row = 18
						exit do
					Elseif FACI_current_panel = FACI_total_check then
						currently_in_FACI = False
						'MsgBox "currently_in_FACI = False"
						exit do
					Else
						transmit
						'MsgBox "transmit"
					End if
				Loop until FACI_current_panel = FACI_total_check
				'stat_faci_currently_in_facility = currently_in_FACI

				If currently_in_FACI = True then
					EmReadscreen psn_faci_vendor_number, 8, 5, 43
					EMReadScreen psn_faci_vendor_name, 30, 6, 43
					EMReadScreen psn_faci_facility_type, 2, 7, 43
					EMReadScreen psn_faci_fs_elig, 1, 8, 43
					EMReadScreen psn_faci_plan_req, 1, 11, 52
					EMReadScreen psn_faci_cty_app_plac, 1, 12, 52
					EMReadScreen psn_faci_postpay, 1, 13, 71
					EMReadScreen psn_faci_grh_rate, 1 , faci_row, 34
					EMReadScreen psn_faci_date_in, 10, faci_row, 47
					EMReadScreen psn_faci_date_out, 10, faci_row, 	71

					psn_faci_vendor_name = trim(replace(psn_faci_vendor_name, "_", ""))
					psn_faci_vendor_number = trim(replace(psn_faci_vendor_number, "_", ""))
					If psn_faci_date_in = "__ __ ____" Then psn_faci_date_in = ""
					psn_faci_date_in = replace(psn_faci_date_in, " ", "/")
					If psn_faci_date_out = "__ __ ____" Then psn_faci_date_out = ""
					psn_faci_date_out = replace(psn_faci_date_out, " ", "/")
					
					'List of FACI types
					If psn_faci_facility_type = "41" Then psn_faci_facility_type = "41-NF-I"
					If psn_faci_facility_type = "42" Then psn_faci_facility_type = "42-NF-II"
					If psn_faci_facility_type = "43" Then psn_faci_facility_type = "43-ICF-DD"
					If psn_faci_facility_type = "44" Then psn_faci_facility_type = "44-Short Stay in NF-I"
					If psn_faci_facility_type = "45" Then psn_faci_facility_type = "45-Short Stay in NF-II"
					If psn_faci_facility_type = "46" Then psn_faci_facility_type = "46-Short Stay in ICF-DD"
					If psn_faci_facility_type = "47" Then psn_faci_facility_type = "47-RTC - Not IMD"
					If psn_faci_facility_type = "48" Then psn_faci_facility_type = "48-Medidcal Hosp IMD Facilities"
					If psn_faci_facility_type = "49" Then psn_faci_facility_type = "49-MSOP"
					If psn_faci_facility_type = "50" Then psn_faci_facility_type = "50-IMD/RTC"
					If psn_faci_facility_type = "51" Then psn_faci_facility_type = "51-Rule 31 CD-IMD"
					If psn_faci_facility_type = "52" Then psn_faci_facility_type = "52-Rule 36 MI-IMD"
					If psn_faci_facility_type = "53" Then psn_faci_facility_type = "53-IMD Hosp GRH Facilities"
					If psn_faci_facility_type = "55" Then psn_faci_facility_type = "55-Adult Foster Care/Rule 203"
					If psn_faci_facility_type = "56" Then psn_faci_facility_type = "56-GRH (not FC or Rule 36)"
					If psn_faci_facility_type = "57" Then psn_faci_facility_type = "57-Rule 36 MI-Non-IMD Other"
					If psn_faci_facility_type = "60" Then psn_faci_facility_type = "60-Non-GRH"
					If psn_faci_facility_type = "61" Then psn_faci_facility_type = "61-Rule 31 CD-Non-IMD"
					If psn_faci_facility_type = "67" Then psn_faci_facility_type = "67-Family Violence Shltr"
					If psn_faci_facility_type = "68" Then psn_faci_facility_type = "68-County Cor Facility"
					If psn_faci_facility_type = "69" Then psn_faci_facility_type = "69-Non-Cty Adult Correctional"
					'MsgBox psn_faci_vendor_number & "-" & psn_faci_vendor_name & "-" & psn_faci_facility_type & "-" & psn_faci_fs_elig & "-" & psn_faci_plan_req & "-" & psn_faci_cty_app_plac & "-" & psn_faci_postpay & "-" & psn_faci_grh_rate & "-" & psn_faci_date_in & "-" & psn_faci_date_out

				End if
			End If

	End If

Next


'Future Iteration -Capturing count of each form so we can iterate the necessary form dialogs -This works well to have after all of the form selection dialogs. Then it doesn't count weird in the do/loop.
' For form_added = 0 to Ubound(form_type_array, 2)
' 	If form_type_array(form_type_const, form_added) = asset_form_name Then asset_count = asset_count + 1 
' 	If form_type_array(form_type_const, form_added) = atr_form_name Then atr_count = atr_count + 1
' 	If form_type_array(form_type_const, form_added) = arep_form_name Then arep_count = arep_count + 1
' 	If form_type_array(form_type_const, form_added) = change_form_name Then change_count = change_count + 1 
' 	If form_type_array(form_type_const, form_added) = evf_form_name Then evf_count = evf_count + 1  
' 	If form_type_array(form_type_const, form_added) = hosp_form_name Then hosp_count = hosp_count + 1 
' 	If form_type_array(form_type_const, form_added) = iaa_form_name Then iaa_count = iaa_count + 1 
' 	If form_type_array(form_type_const, form_added) = iaa_ssi_form_name Then iaa_ssi_count = iaa_ssi_count + 1 
' 	If form_type_array(form_type_const, form_added) = ltc_1503_form_name Then ltc_1503_count = ltc_1503_count + 1 
' 	If form_type_array(form_type_const, form_added) = mof_form_name Then mof_count = mof_count + 1 
' 	If form_type_array(form_type_const, form_added) = mtaf_form_name Then mtaf_count = mtaf_count + 1 
' 	If form_type_array(form_type_const, form_added) = psn_form_name Then psn_count = psn_count + 1 
' 	If form_type_array(form_type_const, form_added) = sf_form_name Then sf_count = sf_count + 1 
' 	If form_type_array(form_type_const, form_added) = diet_form_name Then diet_count = diet_count + 1 
' Next
' MsgBox "checking count of each form" & vbcr & "Asset count " & asset_count & vbcr & "ATR count " & atr_count & vbcr & "AREP " & arep_count & vbcr & "chng " & change_count & vbcr & "evf " & evf_count & vbcr & "hosp " & hosp_count & vbcr & "iaa " & iaa_count & vbcr & "iaa-ssi " & iaa_ssi_count & vbcr & "ltc-1503 " & ltc_1503_count & vbcr & "mof " & mof_count & vbcr & "mtaf " & mtaf_count & vbcr & "psn " & psn_count & vbcr & "sf " & sf_count & vbcr & "diet " & diet_count	'TEST


'DIALOG DISPLAYING FORM SPECIFIC INFORMATION===========================================================================
'Displays individual dialogs for each form selected via checkbox or dropdown. Do/Loops allows us to jump around/are more flexible than For/Next 
form_count = 0
Do	
	Do
		Do
			err_msg = ""
			asset_err_msg = ""
			atr_err_msg = ""
			arep_err_msg = ""
			chng_err_msg = ""
			evf_err_msg = "" 
			hosp_err_msg = ""
			iaa_err_msg = ""
			iaa_ssi_err_msg = ""
			ltc_1503_err_msg = ""
			mof_err_msg = ""
			mtaf_err_msg = ""
			psn_err_msg = ""
			sf_err_msg = ""
			diet_err_msg = ""

			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 456, 300, "Documents Received"
				If form_type_array(form_type_const, form_count) = asset_form_name then Call asset_dialog
				If form_type_array(form_type_const, form_count) = atr_form_name Then Call atr_dialog
				If form_type_array(form_type_const, form_count) = arep_form_name then Call arep_dialog
				If form_type_array(form_type_const, form_count) = change_form_name Then Call change_dialog
				If form_type_array(form_type_const, form_count) = evf_form_name Then Call evf_dialog
				If form_type_array(form_type_const, form_count) = hosp_form_name Then Call hospice_dialog
				If form_type_array(form_type_const, form_count) = iaa_form_name Then Call iaa_dialog
				If form_type_array(form_type_const, form_count) = iaa_ssi_form_name Then Call iaa_ssi_dialog
				If form_type_array(form_type_const, form_count) = ltc_1503_form_name Then Call ltc_1503_dialog
				If form_type_array(form_type_const, form_count) = mof_form_name Then Call mof_dialog
				If form_type_array(form_type_const, form_count) = mtaf_form_name Then Call mtaf_dialog
				If form_type_array(form_type_const, form_count) = psn_form_name Then Call psn_dialog
				If form_type_array(form_type_const, form_count) = sf_form_name Then Call sf_dialog
				If form_type_array(form_type_const, form_count) = diet_form_name Then Call diet_dialog
				
				btn_pos = 45		'variable to iterate down for each necessary button
				''Future Iteration - handle to uniquely identify multiples of the same form by adding count to the button name
				For current_form = 0 to Ubound(form_type_array, 2) 		'This iterates through the array and creates buttons for each form selected from top down. Also stores button name and number in the array based on form name selected. 
					If form_type_array(form_type_const, current_form) = asset_form_name then
						form_type_array(btn_name_const, form_count) = "ASSET"
						form_type_array(btn_number_const, form_count) = 400
						PushButton 395, btn_pos, 45, 15, "ASSET", asset_btn
						'PushButton 395, btn_pos, 45, 15, "ASSET-" & asset_count, asset_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
						' MsgBox "asset name" & form_type_array(form_type_const, current_form) 'TEST
						' MsgBox "asset btn" & form_type_array(btn_name_const, form_count)	'TEST
						' MsgBox "asset numb" & form_type_array(btn_number_const, form_count) 'TEST
					End If
					If form_type_array(form_type_const, current_form) = atr_form_name Then 
						form_type_array(btn_name_const, form_count) = "ATR"
						form_type_array(btn_number_const, form_count) = 401
						PushButton 395, btn_pos, 45, 15, "ATR", atr_btn
						'PushButton 395, btn_pos, 45, 15, "ATR-" & atr_count, atr_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					' 	MsgBox "atr name" & form_type_array(form_type_const, current_form) 'TEST
					' 	MsgBox "atr btn" & form_type_array(btn_name_const, form_count)	'TEST
					' 	MsgBox "atr numb" & form_type_array(btn_number_const, form_count) 'TEST
					End If
					If form_type_array(form_type_const, current_form) = arep_form_name then 
						form_type_array(btn_name_const, form_count) = "AREP"
						form_type_array(btn_number_const, form_count) = 402
						PushButton 395, btn_pos, 45, 15, "AREP", arep_btn
						'PushButton 395, btn_pos, 45, 15, "AREP-" & arep_count, arep_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = change_form_name  then 
						form_type_array(btn_name_const, form_count) = "CHNG"
						form_type_array(btn_number_const, form_count) = 403
						PushButton 395, btn_pos, 45, 15, "CHNG", change_btn 
						'PushButton 395, btn_pos, 45, 15, "CHNG-" & change_count, change_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = evf_form_name  then 
						form_type_array(btn_name_const, form_count) = "EVF"
						form_type_array(btn_number_const, form_count) = 404		
						PushButton 395, btn_pos, 45, 15, "EVF", evf_btn 
						'PushButton 395, btn_pos, 45, 15, "EVF-" & evf_count, evf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = hosp_form_name  then 
						form_type_array(btn_name_const, form_count) = "HOSP"
						form_type_array(btn_number_const, form_count) = 405
						PushButton 395, btn_pos, 45, 15, "HOSP", hospice_btn 
						'PushButton 395, btn_pos, 45, 15, "HOSP-" & hosp_count, hospice_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = iaa_form_name  then 
						form_type_array(btn_name_const, form_count) = "IAA"
						form_type_array(btn_number_const, form_count) = 406
						PushButton 395, btn_pos, 45, 15, "IAA", iaa_btn
						'PushButton 395, btn_pos, 45, 15, "IAA-" & iaa_count, iaa_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = iaa_ssi_form_name then 
						form_type_array(btn_name_const, form_count) = "IAA-SSI"
						form_type_array(btn_number_const, form_count) = 407
						PushButton 395, btn_pos, 45, 15, "IAA-SSI", iaa_ssi_btn 
						'PushButton 395, btn_pos, 45, 15, "IAA-SSI-" & iaa_ssi_count, iaa_ssi_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = ltc_1503_form_name then 
						form_type_array(btn_name_const, form_count) = ltc_1503_form_name
						form_type_array(btn_number_const, form_count) = 408
						PushButton 395, btn_pos, 45, 15, ltc_1503_form_name, ltc_1503_btn 
						'PushButton 395, btn_pos, 45, 15, "LTC-1503-" & ltc_1503_count, ltc_1503_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = mof_form_name then 
						form_type_array(btn_name_const, form_count) = "MOF"
						form_type_array(btn_number_const, form_count) = 409
						PushButton 395, btn_pos, 45, 15, "MOF", mof_btn 
						'PushButton 395, btn_pos, 45, 15, "MOF-" & mof_count, mof_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = mtaf_form_name then 
						form_type_array(btn_name_const, form_count) = "MTAF"
						form_type_array(btn_number_const, form_count) = 410
						PushButton 395, btn_pos, 45, 15, "MTAF", mtaf_btn
						'PushButton 395, btn_pos, 45, 15, "MTAF-" & mtaf_count, mtaf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = psn_form_name then 
						form_type_array(btn_name_const, form_count) = "PSN"
						form_type_array(btn_number_const, form_count) = 411
						PushButton 395, btn_pos, 45, 15, "PSN", psn_btn 
						'PushButton 395, btn_pos, 45, 15, "PSN-" & psn_count, psn_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = sf_form_name then 
						form_type_array(btn_name_const, form_count) = "SF"
						form_type_array(btn_number_const, form_count) = 412
						PushButton 395, btn_pos, 45, 15, "SF", sf_btn
						'PushButton 395, btn_pos, 45, 15, "SF-" & sf_count, sf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = diet_form_name then 
						form_type_array(btn_name_const, form_count) = "DIET"
						form_type_array(btn_number_const, form_count) = 413
						PushButton 395, btn_pos, 45, 15, "DIET", diet_btn
						'PushButton 395, btn_pos, 45, 15, "DIET-" & diet_count, diet_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					'MsgBox "Current form" & form_type_array(form_type_const, current_form)
				Next

				If form_count > 0 Then PushButton 395, 255, 50, 15, "Previous", previous_btn ' Previous button to navigate from one form to the previous one.
				If form_count < Ubound(form_type_array, 2) Then PushButton 395, 275, 50, 15, "Next Form", next_btn	'Next button to navigate from one form to the next. 
				If form_count = Ubound(form_type_array, 2) Then PushButton 395, 275, 50, 15, "Complete", complete_btn	'Complete button kicks off the casenoting of all completed forms. 
				'MsgBox "Ubound(form_type_array, 2)" & Ubound(form_type_array, 2) 
								
			EndDialog
			dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
			cancel_confirmation
			
			
			Call dialog_movement	'function to move throughout the dialogs
			Call main_error_handling	'function for error handling of main dialog of forms 

		Loop until err_msg = ""
	
		'MsgBox "complete_btn" & complete_btn
	Loop until ButtonPressed = complete_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
'MsgBox "Date Effective: " & chng_effective_date + vbCr + "Date Received" & chng_date_received + vbCr + "Address" & chng_address_notes + vbCr + "Household Members" & chng_household_notes + vbCr + "Assets" & chng_asset_notes + vbCr + "Vehicles" & chng_vehicles_notes + vbCr + "Income" & chng_income_notes + vbCr + "Shelter" & chng_shelter_notes + vbCr + "Other" & chng_other_change_notes + vbCr + "Action Taken" & chng_actions_taken + vbCr + "Other Notes" & chng_other_notes + vbCr + "Verifs Requested" & chng_verifs_requested + vbCr + "The changes client reports" & chng_changes_continue		'TEST

'WRITE IN MAXIS===========================================================================
Call MAXIS_background_check
For maxis_panel_write = 0 to Ubound(form_type_array, 2)	
	If form_type_array(form_type_const, maxis_panel_write) = arep_form_name then 	' WRITE FOR AREP
		end_msg = end_msg & vbNewLine & "AREP Information entered."
		'formatting programs into one variable to write in case note
		IF arep_SNAP_AREP_checkbox = checked THEN AREP_programs = "SNAP"
		IF arep_HC_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", HC"
		IF arep_CASH_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", CASH"
		If left(AREP_programs, 1) = "," Then AREP_programs = right(AREP_programs, len(AREP_programs)-2)

		docs_rec = docs_rec & ", AREP Form"

		If arep_update_AREP_panel_checkbox = checked Then				'If update AREP checkbox is checked, then update panel 
			Call MAXIS_background_check
			If IsDate(arep_signature_date) = TRUE Then					'Determing footermonth 
				Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, arep_signature_date)
			Else
				Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, AREP_recvd)
			End If
			Call back_to_SELF
			Do
				Call navigate_to_MAXIS_screen("STAT", "AREP")		'Navigate to AREP panel
				EMReadScreen panel_check, 4, 2, 53
				EMWaitReady 0, 0
			Loop until panel_check = "AREP"

			EMReadScreen arep_version, 1, 2, 73
			If arep_version = "1" Then PF9
			If arep_version = "0" Then Call write_value_and_transmit("NN", 20, 79)

			'Writing to the panel
			EMWriteScreen "                                     ", 4, 32
			EMWriteScreen "                      ", 5, 32
			EMWriteScreen "                      ", 6, 32
			EMWriteScreen "               ", 7, 32
			EMWriteScreen "  ", 7, 55
			EMWriteScreen "     ", 7, 64

			MsgBox "street" & arep_street

			EMWriteScreen arep_name, 4, 32
			arep_street = trim(arep_street)
			If len(arep_street) > 22 Then
				arep_street_one = ""
				arep_street_two = ""
				street_array = split(arep_street, " ")
				For each word in street_array
					If len(arep_street_one & word) > 22 Then
						arep_street_two = arep_street_two & word & " "
					Else
						arep_street_one = arep_street_one & word & " "
					End If
				Next
			Else
				arep_street_one = arep_street
			End If
			EMWriteScreen arep_street_one, 5, 32
			EMWriteScreen arep_street_two, 6, 32
			EMWriteScreen arep_city, 7, 32
			EMWriteScreen arep_state, 7, 55
			EMWriteScreen arep_zip, 7, 64
			EMWriteScreen "N", 5, 77

			If arep_phone_one <> "" Then
				write_phone_one = replace(arep_phone_one, "(", "")
				write_phone_one = replace(write_phone_one, ")", "")
				write_phone_one = replace(write_phone_one, "-", "")
				write_phone_one = trim(write_phone_one)

				EMWriteScreen left(write_phone_one, 3), 8, 34
				EMwriteScreen right(left(write_phone_one, 6), 3), 8, 40
				EMWriteScreen right(write_phone_one, 4), 8, 44
				If arep_ext_one = "" Then
					EMWriteScreen "   ", 8, 55
				Else
					EMWriteScreen arep_ext_one, 8, 55
				End If
			End If

			If arep_phone_two <> "" Then
				write_phone_two = replace(arep_phone_two, "(", "")
				write_phone_two = replace(write_phone_two, ")", "")
				write_phone_two = replace(write_phone_two, "-", "")
				write_phone_two = trim(write_phone_two)

				EMWriteScreen left(write_phone_two, 3), 9, 34
				EMwriteScreen right(left(write_phone_two, 6), 3), 9, 40
				EMWriteScreen right(write_phone_two, 4), 9, 44
				If arep_ext_two = "" Then
					EMWriteScreen "   ", 9, 55
				Else
					EMWriteScreen arep_ext_two, 9, 55
				End If
			End If

			If arep_forms_to_arep_checkbox = checked Then EMWriteScreen "Y", 10, 45
			If arep_forms_to_arep_checkbox = unchecked Then EMWriteScreen "N", 10, 45
			If arep_mmis_mail_to_arep_checkbox = checked Then EMWriteScreen "Y", 10, 77
			If arep_mmis_mail_to_arep_checkbox = unchecked Then EMWriteScreen "N", 10, 77

			transmit
		End If
    End If

	If form_type_array(form_type_const, maxis_panel_write) = iaa_form_name Then	'MAXIS NAVIGATION FOR IAA read/write SCREEN		'TODO: Do we need this for IAA form and IAA-SSI form? Should there be different options for the IAA form?
		If iaa_update_pben_checkbox = checked Then 
			Do
				Call Navigate_to_MAXIS_screen ("STAT", "PBEN")					'Go to PBEN 
				EMReadScreen nav_check, 4, 2, 49
				EMWaitReady 0, 0
			Loop until nav_check = "PBEN"
			
			iaa_referral_date_month = right("00" & DatePart("m", iaa_referral_date), 2)		'Setting up the parts of the date for MAXIS fields
			iaa_referral_date_day = right("00" & DatePart("d", iaa_referral_date), 2)
			iaa_referral_date_year = right(DatePart("yyyy", iaa_referral_date), 2)
			
			iaa_date_applied_pben_month = right("00" & DatePart("m", iaa_date_applied_pben), 2)		'Setting up the parts of the date for MaXIS fields	
			iaa_date_applied_pben_day = right("00" & DatePart("d", iaa_date_applied_pben), 2)
			iaa_date_applied_pben_year = right(DatePart("yyyy", iaa_date_applied_pben), 2)
			
			iaa_date_month = right("00" & DatePart("m", iaa_date), 2)		'Setting up the parts of the date for MAXIS fields	
			iaa_date_day = right("00" & DatePart("d", iaa_date), 2)
			iaa_date_year = right(DatePart("yyyy", iaa_date), 2)
			pben_member_number = Left(iaa_member_dropdown, 2)
			
			Call write_value_and_transmit(pben_member_number, 20, 76)			'Go to the correct member 
			
			pben_row = 8
			'pben_disp_code_string = "*"

			Do 
				EMReadScreen pben_exist, 2, pben_row, 24
				If pben_exist = "__" Then 										
					EMReadScreen numb_of_panels, 1, 2, 78
					IF numb_of_panels = "0" Then 										'If PBEN panel does not exist, create a panel, write dialog entries into fields
						Call write_value_and_transmit("NN", 20, 79)								
					Else
						PF9																'If PBEN panel exists but benefit type is empty, write dialog entries into fields
					End IF
					EMWaitReady 0, 0
					EMWriteScreen Left(iaa_benefit_type, 2), pben_row, 24				'Filling out the panel
					EMWriteScreen iaa_referral_date_month, pben_row, 40
					EMWriteScreen iaa_referral_date_day, pben_row, 43
					EMWriteScreen iaa_referral_date_year, pben_row, 46
					EMWriteScreen iaa_date_applied_pben_month, pben_row, 51
					EMWriteScreen iaa_date_applied_pben_day, pben_row, 54
					EMWriteScreen iaa_date_applied_pben_year, pben_row, 57
					EMWriteScreen Left(iaa_verification_dropdown, 1), pben_row, 62
					EMWriteScreen iaa_date_month, pben_row, 66
					EMWriteScreen iaa_date_day, pben_row, 69
					EMWriteScreen iaa_date_year, pben_row, 72
					EMWriteScreen Left(iaa_disposition_code_dropdown, 1), pben_row, 77
					Exit Do

				ElseIf pben_exist = "02" Then 
					If Left(iaa_benefit_type, 2) = "02" Then 								'If 02 benefit type already exists, must evaluate to see if it is AEPN status. If so, we cannot update the panel. 
						MsgBox "benefit type" & Left(iaa_benefit_type, 2)					'Read line of code if benefit type is 02
						EMReadScreen pben_benefit_type, 2, pben_row, 24
						EMReadScreen pben_referral_date, 8, pben_row, 40
						EMReadScreen pben_date_applied, 8, pben_row, 51
						EMReadScreen pben_verification, 1, pben_row, 62
						EMReadScreen pben_iaa_date, 8, pben_row, 66
						EMReadScreen pben_disp_code, 1, pben_row, 77
						pben_disp_code_string = pben_disp_code_string & pben_disp_code		
				
						If Instr(pben_disp_code_string, "A") or Instr(pben_disp_code_string, "E") or Instr(pben_disp_code_string, "P") or Instr(pben_disp_code_string, "N") Then 
							MsgBox "instring contains AEPN"
							If Left(iaa_disposition_code_dropdown, 1) = "A" or Left(iaa_disposition_code_dropdown, 1) = "E" or Left(iaa_disposition_code_dropdown, 1) = "P" or Left(iaa_disposition_code_dropdown, 1) = "N" Then 
								MsgBox "Cannot update pben panel because there is already an SSI entry with an active disposition code. Manually update PBEN after the script run."
								iaa_update_pben_checkbox = unchecked
								Exit Do 
							Else 
								pben_row = pben_row + 1
							End If
						Else 
							pben_row = pben_row + 1
						End IF 
					Else 
						pben_row = pben_row + 1
					End If
				Else 
					pben_row = pben_row + 1
				End If
			Loop Until pben_row = 14
			If pben_row = 14 Then  				'If all lines on the panel are full then it cannot update PBEN
				MsgBox "PBEN panel is full. Script cannot updated PBEN automatically. Manually update it after script run."
				iaa_update_pben_checkbox = unchecked
			End If
		End If
	End If 


	'TODO: Do we want to look in SVES to verify SSI application date? FOR IAA-SSI and maybe IAA form too? Potentially use function send_sves_qury(ssn_or_claim, qury_finish)
	If form_type_array(form_type_const, maxis_panel_write) = iaa_ssi_form_name Then	'MAXIS NAVIGATION FOR IAA-SSI read/write SCREEN
		If iaa_ssi_update_pben_checkbox = checked Then 
			Do
				Call Navigate_to_MAXIS_screen ("STAT", "PBEN")					'Go to PBEN 
				EMReadScreen nav_check, 4, 2, 49
				EMWaitReady 0, 0
			Loop until nav_check = "PBEN"

			iaa_ssi_referral_date_month = right("00" & DatePart("m", iaa_ssi_referral_date), 2)		'Setting up the parts of the date for MAXIS fields
			iaa_ssi_referral_date_day = right("00" & DatePart("d", iaa_ssi_referral_date), 2)
			iaa_ssi_referral_date_year = right(DatePart("yyyy", iaa_ssi_referral_date), 2)
			
			iaa_ssi_date_applied_pben_month = right("00" & DatePart("m", iaa_ssi_date_applied_pben), 2)		'Setting up the parts of the date for MaXIS fields	
			iaa_ssi_date_applied_pben_day = right("00" & DatePart("d", iaa_ssi_date_applied_pben), 2)
			iaa_ssi_date_applied_pben_year = right(DatePart("yyyy", iaa_ssi_date_applied_pben), 2)
			
			iaa_ssi_iaa_date_month = right("00" & DatePart("m", iaa_ssi_iaa_date), 2)		'Setting up the parts of the date for MAXIS fields	
			iaa_ssi_iaa_date_day = right("00" & DatePart("d", iaa_ssi_iaa_date), 2)
			iaa_ssi_iaa_date_year = right(DatePart("yyyy", iaa_ssi_iaa_date), 2)
			pben_ssi_member_number = Left(iaa_ssi_member_dropdown, 2)
			
			Call write_value_and_transmit(pben_ssi_member_number, 20, 76)			'Go to the correct member 
			
			pben_row = 8

			Do 
				EMReadScreen pben_exist, 2, pben_row, 24
				If pben_exist = "__" Then 										
					EMReadScreen numb_of_panels, 1, 2, 78
					IF numb_of_panels = "0" Then 						'If PBEN panel does not exist, create a panel, write dialog entries into fields
						Call write_value_and_transmit("NN", 20, 79)								
					Else
						PF9												'If PBEN panel exists but benefit type is empty, write dialog entries into fields
					End IF
					EMWaitReady 0, 0
					EMWriteScreen Left(iaa_ssi_benefit_type, 2), pben_row, 24				'Filling out the panel
					EMWriteScreen iaa_ssi_referral_date_month, pben_row, 40
					EMWriteScreen iaa_ssi_referral_date_day, pben_row, 43
					EMWriteScreen iaa_ssi_referral_date_year, pben_row, 46
					EMWriteScreen iaa_ssi_date_applied_pben_month, pben_row, 51
					EMWriteScreen iaa_ssi_date_applied_pben_day, pben_row, 54
					EMWriteScreen iaa_ssi_date_applied_pben_year, pben_row, 57
					EMWriteScreen Left(iaa_ssi_verification_dropdown, 1), pben_row, 62
					EMWriteScreen iaa_ssi_iaa_date_month, pben_row, 66
					EMWriteScreen iaa_ssi_iaa_date_day, pben_row, 69
					EMWriteScreen iaa_ssi_iaa_date_year, pben_row, 72
					EMWriteScreen Left(iaa_ssi_disposition_code_dropdown, 1), pben_row, 77
					Exit Do
					
				ElseIf pben_exist = "02" Then 								'If 02 benefit type already exists, must evaluate to see if it is AEPN status. If so, we cannot update the panel. 
					If Left(iaa_ssi_benefit_type, 2) = "02" Then		'Read line of code if benefit type is 02
						MsgBox "benefit type" & Left(iaa_ssi_benefit_type, 2)
						EMReadScreen pben_benefit_type, 2, pben_row, 24
						EMReadScreen pben_referral_date, 8, pben_row, 40
						EMReadScreen pben_date_applied, 8, pben_row, 51
						EMReadScreen pben_verification, 1, pben_row, 62
						EMReadScreen pben_iaa_date, 8, pben_row, 66
						EMReadScreen pben_disp_code, 1, pben_row, 77
						pben_ssi_disp_code_string = pben_ssi_disp_code_string & pben_disp_code		
				
						If Instr(pben_ssi_disp_code_string, "A") or Instr(pben_ssi_disp_code_string, "E") or Instr(pben_ssi_disp_code_string, "P") or Instr(pben_ssi_disp_code_string, "N") Then 		'Cannot fill out PBEN if 02 exists with AEPN status already. 
							MsgBox "instring contains AEPN"
							If Left(iaa_ssi_disposition_code_dropdown, 1) = "A" or Left(iaa_ssi_disposition_code_dropdown, 1) = "E" or Left(iaa_ssi_disposition_code_dropdown, 1) = "P" or Left(iaa_ssi_disposition_code_dropdown, 1) = "N" Then 
								MsgBox "Cannot update pben panel because there is already an SSI entry with an active disposition code. Manually update PBEN after the script run."
								iaa_ssi_update_pben_checkbox = unchecked
								Exit Do 
							Else 
								pben_row = pben_row + 1
							End If
						Else 
							pben_row = pben_row + 1
						End IF 
					Else 
						pben_row = pben_row + 1
					End If
				Else 
					pben_row = pben_row + 1
				End If
			Loop Until pben_row = 14
			If pben_row = 14 Then 				'If all lines on the panel are full then it cannot update PBEN
				MsgBox "PBEN panel is full. Script cannot updated PBEN automatically. Manually update it after script run."				
				iaa_ssi_update_pben_checkbox = unchecked
			End IF
		End If
	End If 


	If form_type_array(form_type_const, maxis_panel_write) = ltc_1503_form_name then 	' WRITE FOR LTC 1503
		end_msg = end_msg & vbNewLine & "LTC 1503 Form information entered."
		Original_footer_month = MAXIS_footer_month
		Original_footer_year = MAXIS_footer_year
		MAXIS_footer_month = ltc_1503_faci_footer_month
		MAXIS_footer_year = ltc_1503_faci_footer_year
		'LTC 1503 gets it's own case note
		'navigating the script to the correct footer month
		back_to_self
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		Do
			Call navigate_to_MAXIS_screen("STAT", "FACI")
			EMReadScreen nav_check, 4, 2, 44
			EMWaitReady 0, 0
		Loop until nav_check = "FACI"
		MsgBox "are we in faci?"

		'FACI
		If ltc_1503_FACI_update_checkbox = checked then		'If update FACI checkbox checked udpate panel
			Do
				call navigate_to_MAXIS_screen("STAT", "FACI")	'Navigate to FACI
				EMReadScreen nav_check, 4, 2, 44
				EMWaitReady 0, 0
			Loop until nav_check = "FACI"
			EMReadScreen panel_max_check, 1, 2, 78
			IF panel_max_check = "5" THEN			'If panel has reached 5 which is the max, it will not update
				stop_or_continue = MsgBox("This case has reached the maxzimum amount of FACI panels. Please review the case and delete an appropriate FACI panel." & vbNewLine & vbNewLine & "To continue the script run without updating FACI, press 'OK'." & vbNewLine & vbNewLine & "Otherwise, press 'CANCEL' to stop the script, and then rerun it with fewer than 5 FACI panels.", vbQuestion + vbOkCancel, "Continue without updating FACI?")
				If stop_or_continue = vbCancel Then script_end_procedure("~PT User Pressed Cancel")
				If stop_or_continue = vbOk Then ltc_1503_FACI_update_checkbox = unchecked
			ELSE										'Else, create a new panel
				EMWriteScreen "NN", 20, 79
				transmit
			END IF
		End If
		If ltc_1503_FACI_update_checkbox = checked then		'If update FACI checkbox checked udpate panel
			updated_FACI_checkbox = checked
			EMWriteScreen ltc_1503_FACI_1503, 6, 43
			If ltc_1503_level_of_care = "NF" then EMWriteScreen "42", 7, 43
			If ltc_1503_level_of_care = "RTC" THEN EMWriteScreen "47", 7, 43
			If ltc_1503_length_of_stay = "30 days or less" and ltc_1503_level_of_care = "SNF" then EMWriteScreen "44", 7, 43
			If ltc_1503_length_of_stay = "31 to 90 days" and ltc_1503_level_of_care = "SNF" then EMWriteScreen "41", 7, 43
			If ltc_1503_length_of_stay = "91 to 180 days" and ltc_1503_level_of_care = "SNF" then EMWriteScreen "41", 7, 43
			if ltc_1503_length_of_stay = "over 180 days" and ltc_1503_level_of_care = "SNF" then EMWriteScreen "41", 7, 43
			If ltc_1503_length_of_stay = "30 days or less" and ltc_1503_level_of_care = "ICF-DD" then EMWriteScreen "44", 7, 43
			If ltc_1503_length_of_stay = "31 to 90 days" and ltc_1503_level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
			If ltc_1503_length_of_stay = "91 to 180 days" and ltc_1503_level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
			If ltc_1503_length_of_stay = "over 180 days" and ltc_1503_level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
			EMWriteScreen "N", 8, 43
			Call create_MAXIS_friendly_date_with_YYYY(ltc_1503_admit_date, 0, 14, 47)
			If ltc_1503_discharge_date <> "" then
				Call create_MAXIS_friendly_date_with_YYYY(ltc_1503_discharge_date, 0, 14, 71)
				transmit
				transmit
			End if
		End if

		'HCMI
		If ltc_1503_HCMI_update_checkbox = checked THEN
			Do
				Call navigate_to_MAXIS_screen("STAT", "HCMI")
				EMReadScreen nav_check, 4, 2, 55
				EMWaitReady 0, 0
			Loop until nav_check = "HCMI"
			EMReadScreen HCMI_panel_check, 1, 2, 78
			IF HCMI_panel_check <> "0" Then
				PF9
			ELSE
				EMWriteScreen "NN", 20, 79
				transmit
			END IF
			EMWriteScreen "DP", 10, 57
			transmit
			transmit
		END IF

		
	'If form_type_array(form_type_const, maxis_panel_write) = psn_form_name then 	' TODO: WRITE FOR PSN	



	End If

		


	If form_type_array(form_type_const, maxis_panel_write) = diet_form_name Then	'Write for DIET form

		If diet_status_dropdown = "Approved" Then			'Only if the diet is approved should we update the pben panel	TODO: is this correct? 
			If diet_mfip_msa_status <> "Not Active/Pending" Then		'Only if the determine program case status determines the case is active or pending on MSA or MFIP will it fill out the DIET panel. 
				back_to_self
				Do
					Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, diet_date_received)		'Identifying Month/Year from date the diet form was received 'TODO: is this the correct way to get the correct footer month? 
					EMWriteScreen MAXIS_footer_month, 20, 43
					EMWriteScreen MAXIS_footer_year, 20, 46
					MsgBox "doc rec" & diet_date_received
					MsgBox "footermonth" & MAXIS_footer_month & MAXIS_footer_year
					Call Navigate_to_MAXIS_screen ("STAT", "DIET")					'Go to DIET 
					EMReadScreen nav_check, 4, 2, 48
					EMWaitReady 0, 0
				Loop until nav_check = "DIET"
				diet_member_number = Left(diet_member_number, 2)					'Grabbing member number from the member dropdown selection
				Call write_value_and_transmit(diet_member_number, 20, 76)			'Go to the correct member 
				
				'DIET_row = 8 'Setting this variable for the next do...loop
				
				EMReadScreen DIET_total, 1, 2, 78
				If DIET_total = 0 then 								'If panel count is 0, then create a panel
					Call write_value_and_transmit("NN", 20, 79)		
				Else												'If panel exists, edit mode, delete panel, create new panel
					PF9
					Call write_value_and_transmit("DEL", 20, 71)	'TODO: Can I just delete the panel before entering the new information? 
					EMWaitReady 0, 0
					Call write_value_and_transmit("NN", 20, 79)	
				End If
				MsgBox "status" & diet_mfip_msa_status
				If diet_mfip_msa_status = "MFIP-Active" or diet_mfip_msa_status = "MFIP-Pending" Then		'If MFIP then write in diet, hard coded
					EMWriteScreen left(diet_1_dropdown, 2), 8, 40
					EMWriteScreen left(diet_verif_1_dropdown, 1), 8, 51
					EMWriteScreen left(diet_2_dropdown, 2), 9, 40
					EMWriteScreen left(diet_verif_2_dropdown, 1), 9, 51		
				ElseIf diet_mfip_msa_status = "MSA-Active" or diet_mfip_msa_status = "MSA-Pending" Then 	'If MSA then write in diets, hard coded
					EMWriteScreen left(diet_1_dropdown, 2), 11, 40
					EMWriteScreen left(diet_verif_1_dropdown, 1), 11, 51
					EMWriteScreen left(diet_2_dropdown, 2), 12, 40
					EMWriteScreen left(diet_verif_2_dropdown, 1), 12, 51
					EMWriteScreen left(diet_3_dropdown, 2), 13, 40
					EMWriteScreen left(diet_verif_3_dropdown, 1), 13, 51
					EMWriteScreen left(diet_4_dropdown, 2), 14, 40
					EMWriteScreen left(diet_verif_4_dropdown, 1), 14, 51
					EMWriteScreen left(diet_5_dropdown, 2), 15, 40
					EMWriteScreen left(diet_verif_5_dropdown, 1), 15, 51
					EMWriteScreen left(diet_6_dropdown, 2), 16, 40
					EMWriteScreen left(diet_verif_6_dropdown, 1), 16, 51
					EMWriteScreen left(diet_7_dropdown, 2), 17, 40
					EMWriteScreen left(diet_verif_7_dropdown, 1), 17, 51
					EMWriteScreen left(diet_8_dropdown, 2), 18, 40
					EMWriteScreen left(diet_verif_8_dropdown, 1), 18, 51
					MsgBox "review panel"
				End If
			End If
		End If
		If diet_status_dropdown = "Denied" Then			'Only if the diet is denied should we delete the pben panel
			back_to_self
			Do
				Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, diet_date_received)		'Identifying Month/Year from date the diet form was received 'TODO: is this the correct way to get the correct footer month? 
				EMWriteScreen MAXIS_footer_month, 20, 43
				EMWriteScreen MAXIS_footer_year, 20, 46
				MsgBox "doc rec" & diet_date_received
				MsgBox "footermonth" & MAXIS_footer_month & MAXIS_footer_year
				Call Navigate_to_MAXIS_screen ("STAT", "DIET")					'Go to DIET 
				EMReadScreen nav_check, 4, 2, 48
				EMWaitReady 0, 0
			Loop until nav_check = "DIET"
			diet_member_number = Left(diet_member_number, 2)					'Grabbing member number from the member dropdown selection
			Call write_value_and_transmit(diet_member_number, 20, 76)			'Go to the correct member 
			EMReadScreen DIET_total, 1, 2, 78
			If DIET_total <> 0 then 								'If panel count is not 0, then delete a panel. (No action to take if the panel count is = 0)
				PF9
				Call write_value_and_transmit("DEL", 20, 71)	'Delete Diet Panel
			End If
		End If
	End if

Next


'CASE NOTE===========================================================================
'TODO- Hospice: Must keep the same header otherwise reading of past case notes won't work/continue -explore how to create separate case notes for each form
		'Call write_variable_in_CASE_NOTE("*** HOSPICE TRANSACTION FORM RECEIVED ***")
		

Call MAXIS_background_check
'Asset Statement Case Notes
If form_type_array(form_type_const, form_count) = asset_form_name then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** ASSET STATEMENT RECEIVED ***")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If
' 'ATR Case Notes
If form_type_array(form_type_const, form_count) = atr_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** ATR RECEIVED FOR" & atr_name & " ***")
	CALL write_bullet_and_variable_in_case_note("Effective Date", atr_effective_date)
	CALL write_bullet_and_variable_in_case_note("Date Received", atr_date_received)
	CALL write_bullet_and_variable_in_case_note("Member", atr_member_dropdown)
	CALL write_bullet_and_variable_in_case_note("Start Date", atr_start_date)
	CALL write_bullet_and_variable_in_case_note("End Date", atr_end_date)
	CALL write_bullet_and_variable_in_case_note("Authorization Type", atr_authorization_type)
	CALL write_bullet_and_variable_in_case_note("Contact Type", atr_contact_type)
	CALL write_bullet_and_variable_in_case_note("  Contact Name", atr_name)
	CALL write_bullet_and_variable_in_case_note("  Address", atr_address)
	CALL write_bullet_and_variable_in_case_note("  City", atr_city)
	CALL write_bullet_and_variable_in_case_note("  State", atr_state)
	CALL write_bullet_and_variable_in_case_note("  Zip Code", atr_zipcode)
	CALL write_bullet_and_variable_in_case_note("  Phone Number", atr_phone_number)

	If atr_eval_treat_checkbox = checked Then
		CALL write_variable_in_case_note("* Record requested will be used to continue evaluation or treatment")
	End If
	If atr_coor_serv_checkbox = checked Then
		CALL write_variable_in_case_note("* Record requested will be used to coordinate services")
	End If
	If atr_elig_serv_checkbox = checked Then
		CALL write_variable_in_case_note("* Record requested will be used to determine eligibility for assistance/service")
	End If
	If atr_court_checkbox = checked Then
		CALL write_variable_in_case_note("* Record requested will be used for court proceedings")
	End If
	If atr_other_checkbox = checked Then
		CALL write_bullet_and_variable_in_case_note("Record requested will be used", atr_other)
	End If
	CALL write_bullet_and_variable_in_case_note("Comments", atr_comments)
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'AREP Case Notes
If form_type_array(form_type_const, form_count) = arep_form_name then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** AREP Received ***")
    call write_variable_in_CASE_NOTE("* Received: " & AREP_recvd_date & ". AREP: " & arep_name)
    If arep_dhs_3437_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the DHS 3437 - MHCP AUTHORIZED REPRESENTATIVE REQUEST Form.")
    If arep_HC_12729_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the HC 12729 - AUTHORIZED REPRESENTATIVE REQUEST Form.")
    If arep_D405_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("  - AREP name on the SNAP AUTHORIZED REPRESENTATIVE CHOICE D405 Form.")
        Call write_variable_in_CASE_NOTE("  - AREP also authorixed to get and use EBT Card.")
    End If
    If arep_CAF_AREP_page_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in the CAF.")
    If arep_HCAPP_AREP_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in a Health Care Application.")
    If arep_power_of_attorney_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP has Power of Attorney Designation.")
    If AREP_programs <> "" then call write_variable_in_CASE_NOTE("  - Programs Authorized for: " & AREP_programs)
    If arep_signature_date <> "" Then call write_variable_in_CASE_NOTE("  - AREP valid start date: " & arep_signature_date)
    Call write_variable_in_CASE_NOTE("  - Client and AREP signed AREP form.")
    IF AREP_ID_check = checked THEN write_variable_in_CASE_NOTE("  - AREP ID on file.")
    IF arep_TIKL_check = checked THEN write_variable_in_CASE_NOTE("  - TIKL'd for 12 months to get new HC AREP form.")
    If arep_update_AREP_panel_checkbox = checked Then write_variable_in_CASE_NOTE("  - AREP panel updated.")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
'LTC Related I think
' call write_bullet_and_variable_in_case_note("ADDR", ADDR)
' call write_bullet_and_variable_in_case_note("FACI", FACI)
' call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
' call write_bullet_and_variable_in_case_note("DISA", DISA)


End If
'Change Reported Case Note
If form_type_array(form_type_const, form_count) = change_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** CHANGE REPORT FORM RECEIVED ***")
	CALL write_bullet_and_variable_in_case_note("Change Effective Date", chng_effective_date)
	CALL write_bullet_and_variable_in_case_note("Notable changes reported", chng_notable_change)
	CALL write_bullet_and_variable_in_case_note("Date Received", chng_date_received)
	CALL write_bullet_and_variable_in_case_note("Address", chng_address_notes)
	CALL write_bullet_and_variable_in_case_note("Household Members", chng_household_notes)
	CALL write_bullet_and_variable_in_case_note("Assets", chng_asset_notes)
	CALL write_bullet_and_variable_in_case_note("Vehicles", chng_vehicles_notes)
	CALL write_bullet_and_variable_in_case_note("Income", chng_income_notes)
	CALL write_bullet_and_variable_in_case_note("Shelter", chng_shelter_notes)
	CALL write_bullet_and_variable_in_case_note("Other", chng_other_change_notes)
	CALL write_bullet_and_variable_in_case_note("Action Taken", chng_actions_taken)
	CALL write_bullet_and_variable_in_case_note("Other Notes", chng_other_notes)
	CALL write_bullet_and_variable_in_case_note("Verifs Requested", chng_verifs_requested)
	CALL write_bullet_and_variable_in_case_note("The changes client reports", chng_changes_continue)
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'EVF Case Notes
If form_type_array(form_type_const, form_count) = evf_form_name Then 
	Call start_a_blank_case_note
	Call write_variable_in_case_note("*** EVF FORM RECEIVED ***")

    Call write_variable_in_CASE_NOTE("* EVF received " & evf_date_received & ": " & EVF_status_dropdown & "*")
	evf_ref_numb = left(evf_client, 2)
    docs_rec = docs_rec & ", EVF for M" & evf_ref_numb
    Call write_variable_in_CASE_NOTE("  - Employer Name: " & evf_employer)
  	Call write_variable_in_CASE_NOTE("  - EVF for HH member: " & evf_ref_numb)
    'for additional information needed
    IF evf_info = "yes" then
        Call write_variable_in_CASE_NOTE("  - Additional Info requested: " & evf_info & "-on " & evf_info_date & " by " & evf_request_info)
    	'If EVF_TIKL_checkbox = checked then call write_variable_in_CASE_NOTE("* TIKL'd for 10 day return.")
    Else
        Call write_variable_in_CASE_NOTE("  - No additional information is needed/requested.")
    END IF
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'Hospice Case Notes
If form_type_array(form_type_const, form_count) = hosp_form_name Then 
	Call start_a_blank_case_note
	Call write_variable_in_case_note("*** HOSPICE TRANSACTION FORM RECEIVED ***")
	'Call write_bullet_and_variable_in_CASE_NOTE("Effective Date", hosp_effective_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Client", hosp_resident_name)
	Call write_bullet_and_variable_in_CASE_NOTE("Hospice Name", hosp_name)
	Call write_bullet_and_variable_in_CASE_NOTE("NPI Number", hosp_npi_number)
	Call write_bullet_and_variable_in_CASE_NOTE("Date of Entry", hosp_entry_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Exit Date", hosp_exit_date)
	'Call write_bullet_and_variable_in_MMIS_NOTE("Exit due to", exit_cause)         'This field is not currently in use so commented out - workers are testing, may add it back in
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS updated as of", hosp_mmis_updated_date)
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS not updated due to", hosp_reason_not_updated)
	Call write_bullet_and_variable_in_CASE_NOTE("Notes", hosp_other_notes)
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If
'IAA Case Notes
If form_type_array(form_type_const, form_count) = iaa_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** INTERIM ASSISTANCE AGREEMENT RECEIVED ***")
	CALL write_bullet_and_variable_in_case_note("Effective Date", iaa_effective_date)
	CALL write_bullet_and_variable_in_case_note("Date Received", iaa_date_received)
	CALL write_bullet_and_variable_in_case_note("Household Member", iaa_member_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Assistance Type", iaa_type_assistance)
	CALL write_bullet_and_variable_in_case_note("Other benefits resident may be eligible for", "   " & iaa_benefits_1 & "   " & iaa_benefits_2 & "   " & iaa_benefits_3 & "   " & iaa_benefits_4)
	CALL write_bullet_and_variable_in_CASE_NOTE("Benefit type", iaa_benefit_type)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verification", iaa_verification_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Disposition Code", iaa_disposition_code_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Date Applied PBEN", iaa_date_applied_pben)
	CALL write_bullet_and_variable_in_CASE_NOTE("Referral Date", iaa_referral_date)
	CALL write_bullet_and_variable_in_CASE_NOTE("IAA Date", iaa_date)
	If iaa_not_signed_30_checkbox = checked Then 
		CALL write_variable_in_case_note("* NOT signed within 30 days of receiving Combined Application Form or Change Report Form. IAA forms signed by the applicant/participant more than 30 days before they submit a signed Combined Application Form (CAF) or Change Report Form (if participant is already open on other cash assistance) are invalid.")
	End If
	If iaa_update_pben_checkbox = checked Then 
		CALL write_variable_in_case_note("* PBEN Panel updated")
	End If

	CALL write_bullet_and_variable_in_case_note("Notes", iaa_comments)
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If





'IAA-SSI Case Notes
If form_type_array(form_type_const, form_count) = iaa_ssi_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** INTERIM ASSISTANCE AGREEMENT-SSI RECEIVED ***")
	CALL write_bullet_and_variable_in_case_note("Effective Date", iaa_ssi_effective_date)
	CALL write_bullet_and_variable_in_case_note("Date Received", iaa_ssi_date_received)
	CALL write_bullet_and_variable_in_case_note("Household Member", iaa_ssi_member_dropdown)
	CALL write_bullet_and_variable_in_case_note("Assistance Type", iaa_ssi_type_of_interim_assistance)
	CALL write_bullet_and_variable_in_CASE_NOTE("Benefit type", iaa_ssi_benefit_type)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verification", iaa_ssi_verification_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Disposition Code", iaa_ssi_disposition_code_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Date Applied PBEN", iaa_ssi_date_applied_pben)
	CALL write_bullet_and_variable_in_CASE_NOTE("Referral Date", iaa_ssi_referral_date)
	CALL write_bullet_and_variable_in_CASE_NOTE("IAA Date", iaa_ssi_iaa_date)

	If iaa_ssi_not_signed_30days_checkbox = checked Then 
		CALL write_variable_in_case_note("NOT signed within 30 days of receiving Combined Application Form or Change Report Form. IAA forms signed by the applicant/participant more than 30 days before they submit a signed Combined Application Form (CAF) or Change Report Form (if participant is already open on other cash assistance) are invalid.")
	End If
	If iaa_ssi_update_pben_checkbox = checked Then 
		CALL write_variable_in_case_note("* PBEN Panel updated")
	End If
	CALL write_bullet_and_variable_in_case_note("Notes", iaa_ssi_comments)
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'LTC 1503 Case Notes
If form_type_array(form_type_const, form_count) = ltc_1503_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** LTC-1503 FORM RECEIVED ***")
    If ltc_1503_processed_1503_checkbox = checked then
      	call write_variable_in_CASE_NOTE("***Processed 1503 from " & ltc_1503_FACI_1503 & "***")
    Else
      	call write_variable_in_CASE_NOTE("***Rec'd 1503 from " & ltc_1503_FACI_1503 & ", DID NOT PROCESS***")
    End if
    Call write_bullet_and_variable_in_case_note("Length of stay", ltc_1503_length_of_stay)
    Call write_bullet_and_variable_in_case_note("Recommended level of care", ltc_1503_level_of_care)
    Call write_bullet_and_variable_in_case_note("Admitted from", ltc_1503_admitted_from)
    Call write_bullet_and_variable_in_case_note("Hospital admitted from", ltc_1503_hospital_admitted_from)
    Call write_bullet_and_variable_in_case_note("Admit date", ltc_1503_admit_date)
    Call write_bullet_and_variable_in_case_note("Discharge date", ltc_1503_discharge_date)
    Call write_variable_in_CASE_NOTE("---")
    If ltc_1503_updated_RLVA_checkbox = checked and ltc_1503_updated_FACI_checkbox = checked then
    	Call write_variable_in_CASE_NOTE("* Updated RLVA and FACI.")
    Else
      	If ltc_1503_updated_RLVA_checkbox = checked then Call write_variable_in_case_note("* Updated RLVA.")
      	If ltc_1503_updated_FACI_checkbox = checked then Call write_variable_in_case_note("* Updated FACI.")
    End if
    If ltc_1503_need_3543_checkbox = checked then Call write_variable_in_case_note("* A 3543 is needed for the client.")
    If ltc_1503_need_3531_checkbox = checked then call write_variable_in_CASE_NOTE("* A 3531 is needed for the client.")
    If ltc_1503_need_asset_assessment_checkbox = checked then call write_variable_in_CASE_NOTE("* An asset assessment is needed before a MA-LTC determination can be made.")
    If ltc_1503_sent_3050_checkbox = checked then call write_variable_in_CASE_NOTE("* Sent 3050 back to LTCF.")
    If ltc_1503_sent_5181_checkbox = checked then call write_variable_in_CASE_NOTE("* Sent DHS-5181 to Case Manager.")
    Call write_bullet_and_variable_in_case_note("Verifs needed", ltc_1503_verifs_needed)
    If ltc_1503_sent_verif_request_checkbox = checked then Call write_variable_in_case_note("* Sent verif request to " & ltc_1503_sent_request_to)
    If processed_1503_checkbox = checked then Call write_variable_in_case_note("* Completed & Returned 1503 to LTCF.")
    If ltc_1503_TIKL_checkbox = checked then Call write_variable_in_case_note("TIKL'd for " & ltc_1503_TIKL_multiplier & " days to check length of stay.")
    Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", ltc_1503_mets_case_number)
    Call write_bullet_and_variable_in_case_note("Notes", ltc_1503_notes)
	CALL write_variable_in_case_note("   ")
    Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
    MAXIS_footer_month = Original_footer_month
    MAXIS_footer_year = Original_footer_year
End If

'MOF Case Notes
If form_type_array(form_type_const, form_count) = mof_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** MEDICAL OPINION FORM RECEIVED ***")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'MTAF Case Notes
If form_type_array(form_type_const, form_count) = mtaf_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** MINNESOTA TRANSITION APPLICATION RECEIVED ***")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'PSN Case Notes
If form_type_array(form_type_const, form_count) = psn_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** PROFESSIONAL STATEMENT OF NEED RECEIVED ***")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'SF Case Notes
If form_type_array(form_type_const, form_count) = sf_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** SHELTER FORM RECEIVED ***")
	CALL write_variable_in_case_note("   ")
	Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
End If

'Special Diet Case Notes
If form_type_array(form_type_const, form_count) = diet_form_name Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** SPECIAL DIET FORM RECEIVED ***")	
	CALL write_bullet_and_variable_in_case_note("Date Effective", diet_effective_date)					
	CALL write_bullet_and_variable_in_case_note("Date Received", diet_date_received)	
	CALL write_bullet_and_variable_in_case_note("Member", Left(diet_member_number, 2))							'required
	If diet_1_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 1", diet_1_dropdown & "- " & diet_relationship_1_dropdown)	'required
	If diet_2_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 2", diet_2_dropdown & "- " & diet_relationship_2_dropdown)	'required
	If diet_3_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 3", diet_3_dropdown & "- " & diet_relationship_3_dropdown)	'required
	If diet_4_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 4", diet_4_dropdown & "- " & diet_relationship_4_dropdown)	'required
	If diet_5_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 5", diet_5_dropdown & "- " & diet_relationship_5_dropdown)	'required
	If diet_6_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 6", diet_6_dropdown & "- " & diet_relationship_6_dropdown)	'required
	If diet_7_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 7", diet_7_dropdown & "- " & diet_relationship_7_dropdown)	'required
	If diet_8_dropdown <> "" Then CALL write_bullet_and_variable_in_case_note("  Diet 8", diet_8_dropdown & "- " & diet_relationship_8_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("Last exam date", diet_date_last_exam)
	CALL write_bullet_and_variable_in_case_note("Diet Length", diet_length_diet)							'required
	CALL write_bullet_and_variable_in_case_note("Person following treatment plan", diet_treatment_plan_dropdown)

	If diet_mfip_msa_status = "Not Active/Pending" Then CALL write_variable_in_case_note("* DIET panel NOT updated- case is not active/pending for MSA or MFIP")		'TODO: Inquire if this is the action or rather inaction we want to take
	If diet_status_dropdown = "Incomplete" then
		CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown & "- form returned to client")
	ElseIf  diet_status_dropdown = "Denied" Then
		CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown & "- Deleted DIET panel. The doctor has not indicated an eligible diet need.")
	Else
		CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown)
	End If 
	CALL write_bullet_and_variable_in_case_note("Comments",diet_comments)
	CALL write_variable_in_case_note("   ")
	CALL write_variable_in_case_note("---")
	CALL write_variable_in_case_note(worker_signature)
End If


'change  
	'If we checked to TIKL out, it goes to TIKL and sends a TIKL
	' IF tikl_nav_check = 1 THEN
	' 	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	' 	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	' 	EMSetCursor 9, 3
	' END IF

'EVF 

	' 	'Checks if additional info is yes and the TIKL is checked, sets a TIKL for the return of the info
	' 	If EVF_TIKL_checkbox = checked Then
	' 		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
	' 		Call create_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action.", 10, date, True, TIKL_note_text)
	' 		'Success message
	' 		end_msg = end_msg & vbNewLine & "Additional detail added about EVF." & vbNewLine & "TIKL has been sent for 10 days from now for the additional information requested."
	' 	Else
	' 		end_msg = end_msg & vbNewLine & "Additional detail added about EVF."
	' 	End If
	' End If

'LTC 1503 
    ' If length_of_stay = "30 days or less"   then TIKL_multiplier = 30
    ' If length_of_stay = "31 to 90 days"     then TIKL_multiplier = 90
    ' If length_of_stay = "91 to 180 days"    then TIKL_multiplier = 180
    ' 'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    ' If TIKL_checkbox = checked then Call create_TIKL("Have " & worker_signature & " call " & FACI & " re: length of stay. " & TIKL_multiplier & " days expired.", TIKL_multiplier, admit_date, False, TIKL_note_text)

  
	
script_end_procedure_with_error_report(end_msg)



'ADDING A NEW FORM TO SCRIPT TO DO LIST 
'Define Form Name
'Define Count Var
'Define BTN Var
'Define Dialog Function 
'Dim Variables
'Dialog Movement- BTN and Form Name
'Drop Down Selection Dialog
'Checkbox VAR
'Checkbox Dialog
'Array Form Capture
'Err Handling
'MSGBOX capturing count as verification
'Capture Count
'Call Dialog
'Btn Name Display Handling 