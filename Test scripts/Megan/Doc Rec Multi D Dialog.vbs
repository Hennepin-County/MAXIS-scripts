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

'FUNCTIONS DEFINED===========================================================================
function asset_dialog()
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, asset_effective_date
			EditBox 310, 20, 45, 15, asset_date_received
			EditBox 30, 65, 270, 15, asset_Q1
			EditBox 30, 85, 270, 15, asset_Q2
			EditBox 30, 105, 270, 15, asset_Q3
			EditBox 30, 125, 270, 15, asset_Q4
			Text 5, 5, 220, 10, "ASSET STATEMENT"
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
		EditBox 30, 65, 270, 15, arep_Q1
		EditBox 30, 85, 270, 15, arep_Q2
		EditBox 30, 105, 270, 15, arep_Q3
		EditBox 30, 125, 270, 15, arep_Q4		
		Text 5, 5, 220, 10, "AREP (Authorized Rep)"
		Text 125, 25, 50, 10, "Effective Date:"
		Text 15, 70, 10, 10, "Q1"
		Text 245, 25, 60, 10, "Document Date:"
		GroupBox 5, 50, 305, 195, "Responses to form questions captured here"
		Text 5, 25, 50, 10, "Case Number:"
		Text 15, 110, 10, 10, "Q3"
		Text 15, 130, 15, 10, "Q4"
		Text 15, 90, 15, 10, "Q2"
		Text 15, 150, 15, 10, ""
end function

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
	Text 5, 5, 220, 10, "CHANGE REPORT FORM"
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

'Dimming all the variables because they are defined and set within functions
Dim chng_effective_date, chng_date_received, chng_address_notes, chng_household_notes, chng_asset_notes, chng_vehicles_notes, chng_income_notes, chng_shelter_notes, chng_other_change_notes, chng_actions_taken, chng_other_notes, chng_verifs_requested, chng_changes_continue, chng_notable_change 'Change Reported variables


function evf_dialog()
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, evf_effective_date
			EditBox 310, 20, 45, 15, evf_date_received
			EditBox 30, 65, 270, 15, evf_Q1
			EditBox 30, 85, 270, 15, evf_Q2
			EditBox 30, 105, 270, 15, evf_Q3
			EditBox 30, 125, 270, 15, evf_Q4
			Text 5, 5, 220, 10, "Employment Verification Form (EVF)"
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

function hospice_dialog()
	EditBox 175, 20, 45, 15, hosp_effective_date
	EditBox 310, 20, 45, 15, hosp_date_received		
	DropListBox 100, 45, 165, 15, HH_Memb_DropDown, hosp_resident_name
	EditBox 100, 65, 205, 15, hosp_name
	EditBox 100, 85, 80, 15, hops_npi_number
	EditBox 100, 105, 50, 15, hosp_entry_date
	EditBox 205, 105, 50, 15, hosp_exit_date
	EditBox 100, 125, 50, 15, hosp_mmis_updated_date
	EditBox 30, 160, 275, 15, hosp_reason_not_updated
	EditBox 30, 190, 275, 15, hosp_other_notes
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "TE 02.07.081", hosp_TE0207081_btn
		PushButton 65, 280, 50, 15, "MA-Hospice", hosp_SP_hospice_btn
	Text 5, 5, 220, 10, "HOSPICE TRANSACTION FORM"
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

Dim hosp_effective_date, hosp_date_received, hosp_resident_name, hosp_name, hops_npi_number, hosp_entry_date, hosp_exit_date, hosp_mmis_updated_date, hosp_reason_not_updated, hosp_other_notes, hosp_TE0207081_btn, hosp_SP_hospice_btn

function iaa_dialog()
	EditBox 175, 15, 45, 15, iaa_effective_date
	EditBox 310, 15, 45, 15, iaa_date_received		
	DropListBox 55, 45, 110, 15, HH_Memb_DropDown, iaa_member_dropdown
	DropListBox 265, 45, 60, 15, ""+chr(9)+"Initial claim"+chr(9)+"Post-eligibility", iaa_type_assistance
	CheckBox 30, 70, 295, 15, "Signed within 30 days of receiving Combined Application Form or Change Report Form.", iaa_within_30_checkbox
	CheckBox 30, 85, 310, 15, "NOT signed within 30 days of receiving Combined Application Form or Change Report Form.", iaa_outside_30_checkbox
	EditBox 50, 125, 145, 15, iaa_benefits_1
	EditBox 50, 145, 145, 15, iaa_benefits_2
	EditBox 205, 125, 145, 15, iaa_benefits_3
	EditBox 205, 145, 145, 15, iaa_benefits_4
	EditBox 50, 180, 300, 15, iaa_comments
	ButtonGroup ButtonPressed
		PushButton 5, 280, 95, 15, "IAA Maxis Instructions", iaa_sp_btn
	Text 5, 5, 220, 10, "INTERIM ASSISTANCE AUTHORIZATION"
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, "MAXIS_case_number"
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 20, 50, 30, 10, "Member"
	Text 175, 50, 90, 10, "Type of interim assistance"
	Text 20, 115, 130, 10, "Other benefits you may be eligible for"
	Text 15, 185, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function 

Dim iaa_effective_date, iaa_date_received, iaa_member_dropdown, iaa_type_assistance, iaa_within_30_checkbox, iaa_outside_30_checkbox, iaa_benefits_1, iaa_benefits_2, iaa_benefits_3, iaa_benefits_4, iaa_comments, iaa_sp_btn


function iaa_ssi_dialog()
	EditBox 175, 20, 45, 15, iaa_ssi_effective_date
	EditBox 310, 20, 45, 15, iaa_ssi_date_received		
	DropListBox 55, 45, 110, 15, HH_Memb_DropDown, iaa_ssi_member_dropdown 
	DropListBox 265, 45, 95, 20, ""+chr(9)+"General Assistance (GA)"+chr(9)+"Housing Support (HS)", iaa_ssi_type_assistance
	CheckBox 30, 70, 295, 15, "Signed within 30 days of receiving Combined Application Form or Change Report Form.", iaa_ssi_within_30_checkbox
	CheckBox 30, 85, 310, 15, "NOT signed within 30 days of receiving Combined Application Form or Change Report Form.", iaa_ssi_outside_30_checkbox
	EditBox 55, 105, 300, 15, iaa_ssi_comments
	ButtonGroup ButtonPressed
	PushButton 5, 280, 50, 15, "CM12.12.03", iaa_ssi_CM121203_btn
	PushButton 65, 280, 95, 15, "IAA-SSI Maxis Instructions", iaa_ssi_sp_btn
	Text 5, 5, 220, 10, "INTERIM ASSISTANCE AUTHORIZATION- SSI"
	Text 5, 25, 50, 10, "Case Number:"
	Text 60, 25, 45, 10, MAXIS_case_number
	Text 125, 25, 50, 10, "Effective Date:"
	Text 245, 25, 60, 10, "Document Date:"
	Text 5, 5, 220, 10, "Interim Assistance Authorization- SSI"
	Text 20, 50, 30, 10, "Member"
	Text 175, 50, 90, 10, "Type of interim assistance"
	Text 20, 110, 35, 10, "Comments"
	Text 395, 35, 45, 10, "    --Forms--"
end function

Dim iaa_ssi_effective_date, iaa_ssi_date_received, iaa_ssi_member_dropdown, iaa_ssi_type_assistance, iaa_ssi_within_30_checkbox, iaa_ssi_outside_30_checkbox, iaa_ssi_comments, iaa_ssi_CM121203_btn, iaa_ssi_sp_btn

function ltc_1503_dialog()
			Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, ltc_1503_effective_date
			EditBox 310, 20, 45, 15, ltc_1503_date_received
			EditBox 30, 65, 270, 15, ltc_1503_Q1
			EditBox 30, 85, 270, 15, ltc_1503_Q2
			EditBox 30, 105, 270, 15, ltc_1503_Q3
			EditBox 30, 125, 270, 15, ltc_1503_Q4			
			Text 5, 5, 220, 10, "LTC-1503"
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

function mof_dialog()
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, mof_effective_date
			EditBox 310, 20, 45, 15, mof_date_received
			EditBox 30, 65, 270, 15, mof_Q1
			EditBox 30, 85, 270, 15, mof_Q2
			EditBox 30, 105, 270, 15, mof_Q3
			EditBox 30, 125, 270, 15, mof_Q4			
			Text 5, 5, 220, 10, "Medical Opinion Form (MOF)"
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
			Text 5, 5, 220, 10, "Minnesota Transition Application Form (MTAF)"
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
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, psn_effective_date
			EditBox 310, 20, 45, 15, psn_date_received
			EditBox 30, 65, 270, 15, psn_Q1
			EditBox 30, 85, 270, 15, psn_Q2
			EditBox 30, 105, 270, 15, psn_Q3
			EditBox 30, 125, 270, 15, psn_Q4			
			Text 5, 5, 220, 10, "Professional Statement of Need (PSN)"
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

function sf_dialog()	
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, sf_effective_date
			EditBox 310, 20, 45, 15, sf_date_received
			EditBox 30, 65, 270, 15, sf_Q1
			EditBox 30, 85, 270, 15, sf_Q2
			EditBox 30, 105, 270, 15, sf_Q3
			EditBox 30, 125, 270, 15, sf_Q4			
			Text 5, 5, 220, 10, "Residence and Shelter Expenses Release Form"
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
	DropListBox 50, 35, 120, 15, HH_Memb_DropDown, diet_member_number 'TODO: Need to populate member number here
	EditBox 50, 55, 290, 15, diet_diagnosis
	DropListBox 55, 85, 115, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_1_dropdown
	DropListBox 185, 85, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_1_dropdown
	DropListBox 55, 100, 115, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_2_dropdown
	DropListBox 185, 100, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_2_dropdown
	DropListBox 55, 115, 115, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_3_dropdown
	DropListBox 185, 115, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_3_dropdown
	DropListBox 55, 130, 115, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_4_dropdown
	DropListBox 185, 130, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_4_dropdown
	EditBox 75, 160, 55, 15, diet_date_last_exam
	DropListBox 130, 180, 35, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_treatment_plan_dropdown
	EditBox 270, 180, 55, 15, diet_length_diet
	DropListBox 130, 200, 60, 15, ""+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", diet_status_dropdown		'TODO: Handling for each scenario- each has it's own notification process/steps
	EditBox 50, 220, 290, 15, diet_prognosis
	EditBox 50, 240, 290, 15, diet_comments
	PushButton 5, 280, 80, 15, "CM23.12- Special Diets", diet_link_CM_special_diet
    PushButton 95, 280, 115, 15, "Processing Special Diet Referrals", diet_SP_referrals
	Text 395, 35, 45, 10, "    --Forms--"
	Text 5, 5, 220, 10, "SPECIAL DIET INFORMATION REQUEST (MFIP and MSA)"
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, MAXIS_case_number
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 20, 40, 30, 10, "Member"
	Text 15, 60, 35, 10, "Diagnosis"
	Text 55, 75, 85, 10, "Select applicable diet"
	Text 185, 75, 95, 10, "Relationship between diets"
	Text 30, 85, 20, 10, "Diet 1"
	Text 30, 100, 20, 10, "Diet 2"
	Text 30, 115, 20, 10, "Diet 3"
	Text 30, 130, 20, 10, "Diet 4"
	Text 15, 165, 60, 10, "Date of last exam"
	Text 15, 185, 115, 10, "Is person following treament plan?"
	Text 185, 185, 85, 10, "Length of Prescribed Diet"
	Text 15, 205, 120, 10, "Diet approved, denied, incomplete?"
	Text 15, 225, 35, 10, "Prognosis"
	Text 15, 245, 35, 10, "Comments"
end function

Dim diet_effective_date, diet_date_received, diet_member_number, diet_diagnosis, diet_1_dropdown, diet_2_dropdown, diet_3_dropdown, diet_4_dropdown, diet_relationship_1_dropdown, diet_relationship_2_dropdown, diet_relationship_3_dropdown, diet_relationship_4_dropdown, diet_date_last_exam, diet_treatment_plan_dropdown, diet_status_dropdown, diet_length_diet, diet_prognosis, diet_comments	'Special Diet Variables


function dialog_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs. 
	If ButtonPressed = -1 Then ButtonPressed = next_btn 	'If the enter button is selected the script will handle this as if Next was selected 
	If ButtonPressed = next_btn Then form_count = form_count + 1	'If next is selected, it will iterate to the next form in the array and display this dialog
	If ButtonPressed = previous_btn Then form_count = form_count - 1	'If previous is selected, it will iterate to the previous form in the array and display this dialog
	If ButtonPressed >= 400 Then 'All forms are in the 400 range
		For i = 0 to Ubound(form_type_array, 2) 	'For/Next used to iterate through the array to display the correct dialog
			If ButtonPressed = asset_btn and form_type_array(form_type_const, i) = "Asset Statement" Then form_count = i 
			If ButtonPressed = atr_btn and form_type_array(form_type_const, i) = "Authorization to Release Information (ATR)" Then form_count = i 
			If ButtonPressed = arep_btn and form_type_array(form_type_const, i) = "AREP (Authorized Rep)" Then form_count = i 
			If ButtonPressed = change_btn and form_type_array(form_type_const, i) = "Change Report Form" Then form_count = i 
			If ButtonPressed = evf_btn and form_type_array(form_type_const, i) = "Employment Verification Form (EVF)" Then form_count = i 
			If ButtonPressed = hospice_btn and form_type_array(form_type_const, i) = "Hospice Transaction Form" Then form_count = i 
			If ButtonPressed = iaa_btn and form_type_array(form_type_const, i) = "Interim Assistance Agreement (IAA)" Then form_count = i 
			If ButtonPressed = iaa_ssi_btn and form_type_array(form_type_const, i) = "Interim Assistance Authorization- SSI" Then form_count = i 
			If ButtonPressed = ltc_1503_btn and form_type_array(form_type_const, i) = "LTC-1503" Then form_count = i 
			If ButtonPressed = mof_btn and form_type_array(form_type_const, i) = "Medical Opinion Form (MOF)" Then form_count = i 
			If ButtonPressed = mtaf_btn and form_type_array(form_type_const, i) = "Minnesota Transition Application Form (MTAF)" Then form_count = i 
			If ButtonPressed = psn_btn and form_type_array(form_type_const, i) = "Professional Statement of Need (PSN)" Then form_count = i 
			If ButtonPressed = sf_btn and form_type_array(form_type_const, i) = "Residence and Shelter Expenses Release Form" Then form_count = i 
			If ButtonPressed = diet_btn and form_type_array(form_type_const, i) = "Special Diet Information Request (MFIP and MSA)" Then form_count = i 
		Next
	End If 
end function 

'Check for case number & footer
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


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
				DropListBox 30, 30, 180, 15, ""+chr(9)+"Asset Statement"+chr(9)+"Authorization to Release Information (ATR)"+chr(9)+"AREP (Authorized Rep)"+chr(9)+"Change Report Form"+chr(9)+"Employment Verification Form (EVF)"+chr(9)+"Hospice Transaction Form"+chr(9)+"Interim Assistance Agreement (IAA)"+chr(9)+"Interim Assistance Authorization- SSI"+chr(9)+"LTC-1503"+chr(9)+"Medical Opinion Form (MOF)"+chr(9)+"Minnesota Transition Application Form (MTAF)"+chr(9)+"Professional Statement of Need (PSN)"+chr(9)+"Residence and Shelter Expenses Release Form"+chr(9)+"Special Diet Information Request (MFIP and MSA)", Form_type
				ButtonGroup ButtonPressed
				PushButton 225, 30, 35, 10, "Add", add_button
				PushButton 225, 60, 35, 10, "All Forms", all_forms
				OkButton 205, 215, 40, 15
				CancelButton 255, 215, 40, 15
				PushButton 155, 215, 40, 15, "Clear", clear_button
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
					If instr(all_form_array, "*" & form_string & "*") Then	
						add_to_array = false	'If the string is found in the array, it won't add the form to the array
					Else 
						add_to_array = true 	'If the string is not found in the array, it will add the form to the array
					End If
				End If
			
				If add_to_array = True Then			'Defining the steps to take if the form should be added to the array
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = Form_type		'Storing form name in the array		
					form_count = form_count + 1 	
					all_form_array = all_form_array & form_string & "*" 'Adding form name to form name string
				ElseIF add_to_array = False then 
					false_count = false_count + 1
				End If 
			End If

			
			MsgBox "all form array string" & all_form_array '= split(all_form_array, "*")

			'This work for handling the adding of each form - this allows you to add more than one of each form 
			' If ButtonPressed = add_button and form_type <> "" Then				'If statement to know when to store the information in the array
			' 	ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
			' 	form_type_array(form_type_const, form_count) = Form_type		'Storing form name in the array		
			' 	form_count= form_count + 1 										'incrementing in the array
			' End If
				
			If ButtonPressed = clear_button Then 'Clear button wipes out any selections already made so the user can reselect correct forms.
				ReDim form_type_array(the_last_const, form_count)		
				form_count = 0							'Reset the form count to 0 so that y_pos resets to 95. 
				asset_checkbox = unchecked				'Resetting checkboxes to unchecked
				atr_checkbox = unchecked				'Resetting checkboxes to unchecked
				arep_checkbox = unchecked				'Resetting checkboxes to unchecked
				change_checkbox = unchecked				'Resetting checkboxes to unchecked
				evf_checkbox = unchecked				'Resetting checkboxes to unchecked
				hospice_checkbox = unchecked			'Resetting checkboxes to unchecked
				iaa_checkbox = unchecked				'Resetting checkboxes to unchecked
				iaa_ssi_checkbox = unchecked			'Resetting checkboxes to unchecked
				ltc_1503_checkbox = unchecked			'Resetting checkboxes to unchecked
				mof_checkbox = unchecked				'Resetting checkboxes to unchecked
				mtaf_checkbox = unchecked				'Resetting checkboxes to unchecked
				psn_checkbox = unchecked				'Resetting checkboxes to unchecked
				shelter_checkbox = unchecked			'Resetting checkboxes to unchecked
				diet_checkbox = unchecked				'Resetting checkboxes to unchecked
				form_type = ""							'Resetting dropdown to blank


				'MsgBox "form string" & form_string
				'MsgBox "all form array" & all_form_array
				asset_count 	= 0 
				atr_count 		= 0 
				arep_count 		= 0 
				change_count 	= 0
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
				Form_string = ""
				all_form_array = ""
				MsgBox "all form array" & all_form_array
			'	MsgBox "form type" & form_type 'TEST
				MsgBox "Form selections cleared." & vbNewLine & "-Make new selection."	'Notify end user that entries were cleared.
			End If

			'TODO: Error handling
			' If add_to_array = false Then
			' 	If form_type = "" Then 
			' 		If ButtonPressed <> all_forms Then
			' 			err_msg = err_msg & "-No form selected- select a form name, then select Add"
			' 		End If
			' 	End If 
				
			' 	If form_type <> "" Then
			' 		If ButtonPressed <> clear_button Then err_msg = err_msg & "-Form already added, select a different form"
			' 		If ButtonPressed = clear_button Then err_msg = err_msg & "-Form selections cleared."
			' 	End If 
			' End If

			If form_count = 0 and ButtonPressed = Ok Then err_msg = "-Add forms to process or select cancel to exit script"		'If form_count = 0, then no forms have been added to doc rec to be processed.	
			If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	If ButtonPressed = all_forms Then		'Opens Dialog with checkbox selection for each form
		Do
			Do
				ReDim form_type_array(the_last_const, form_count)		'Resetting any selections already made so the user can reselect correct forms using different format.
				form_type_array(form_type_const, form_count) = Form_type
                form_count = 0							'Resetting the form count to 0 so that y_pos resets to 95. 
				Form_string = ""						'Resetting string to nothing 
				all_form_array = ""						'Resetting list of strings to nothing 

				'Future Iteration - carries values selected from drop down through to checkbox feature
				' If instr(all_form_array, "Asset Statement") Then asset_checkbox = checked 
				' If instr(all_form_array, "Authorization to Release Information (ATR)") Then atr_checkbox = checked 
				' If instr(all_form_array, "AREP (Authorized Rep)") Then arep_checkbox = checked 
				' If instr(all_form_array, "Change Report Form") Then change_checkbox = checked 
				' If instr(all_form_array, "Employment Verification Form (EVF)") Then evf_checkbox = checked 
				' If instr(all_form_array, "Hospice Transaction Form") Then hospice_checkbox = checked 
				' If instr(all_form_array, "Interim Assistance Agreement (IAA)") Then iaa_checkbox = checked 
				' If instr(all_form_array, "Interim Assistance Authorization- SSI") Then iaa_ssi_checkbox = checked 
				' If instr(all_form_array, "LTC-1503") Then ltc_1503_checkbox = checked 
				' If instr(all_form_array, "Medical Opinion Form (MOF)") Then mof_checkbox = checked 
				' If instr(all_form_array, "Minnesota Transition Application Form (MTAF)") Then mtaf_checkbox = checked 
				' If instr(all_form_array, "Professional Statement of Need (PSN)") Then psn_checkbox = checked 
				' If instr(all_form_array, "Residence and Shelter Expenses Release Form") Then shelter_checkbox = checked 
				' If instr(all_form_array, "Special Diet Information Request (MFIP and MSA)") Then diet_checkbox = checked 


				err_msg = ""
				Dialog1 = "" 'Blanking out previous dialog detail
				BeginDialog Dialog1, 0, 0, 196, 200, "Document Selection"
					CheckBox 15, 20, 160, 10, "Asset Statement", asset_checkbox
					CheckBox 15, 30, 160, 10, "Authorization to Release Information (ATR)", atr_checkbox
					CheckBox 15, 40, 160, 10, "AREP (Authorized Rep)", arep_checkbox
					CheckBox 15, 50, 160, 10, "Change Report Form", change_checkbox
					CheckBox 15, 60, 160, 10, "Employment Verification Form (EVF)", evf_checkbox
					CheckBox 15, 70, 160, 10, "Hospice Transaction Form", hospice_checkbox
					CheckBox 15, 80, 160, 10, "Interim Assistance Agreement (IAA)", iaa_checkbox
					CheckBox 15, 90, 160, 10, "Interim Assistance Authorization- SSI", iaa_ssi_checkbox
					CheckBox 15, 100, 160, 10, "LTC-1503", ltc_1503_checkbox
					CheckBox 15, 110, 160, 10, "Medical Opinion Form (MOF)", mof_checkbox
					CheckBox 15, 120, 160, 10, "Minnesota Transition Application Form (MTAF)", mtaf_checkbox
					CheckBox 15, 130, 160, 10, "Professional Statement of Need (PSN)", psn_checkbox
					CheckBox 15, 140, 170, 10, "Residence and Shelter Expenses Release Form", shelter_checkbox
					CheckBox 15, 150, 175, 10, "Special Diet Information Request (MFIP and MSA)", diet_checkbox
					ButtonGroup ButtonPressed
						'PushButton 70, 180, 70, 15, "Review Selections", review_selections
						OkButton 95, 180, 45, 15
						CancelButton 150, 180, 40, 15
					Text 5, 5, 200, 10, "Select documents received, then Ok."
				EndDialog
				dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation

				
				
				'Capturing form name in array based on checkboxes selected 
				If asset_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Asset Statement" 
					form_count= form_count + 1 
				End If

				If atr_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Authorization to Release Information (ATR)"
					form_count= form_count + 1 
				End If

				If arep_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "AREP (Authorized Rep)"
					form_count= form_count + 1 
				End If

				If change_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Change Report Form"
					form_count= form_count + 1 
				End If
				If evf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Employment Verification Form (EVF)"
					form_count= form_count + 1 
				End If
				If hospice_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Hospice Transaction Form"
					form_count= form_count + 1 
				End If
				If iaa_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Interim Assistance Agreement (IAA)"
					form_count= form_count + 1 
				End If
				If iaa_ssi_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Interim Assistance Authorization- SSI"
					form_count= form_count + 1 
				End If
				If ltc_1503_checkbox = checked Then
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "LTC-1503"
					form_count= form_count + 1 
				End If
				If mof_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Medical Opinion Form (MOF)"
					form_count= form_count + 1 
				End If
				If mtaf_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Minnesota Transition Application Form (MTAF)"
					form_count= form_count + 1 
				End If
				If psn_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Professional Statement of Need (PSN)"
					form_count= form_count + 1 
				End If
				If shelter_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Residence and Shelter Expenses Release Form"
					form_count= form_count + 1 
				End If
				If diet_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = "Special Diet Information Request (MFIP and MSA)"
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
'MAXIS NAV: Hospice Read Screen ===========================================================================
For maxis_panel_read = 0 to Ubound(form_type_array, 2)
	If form_type_array(form_type_const, form_added) = "Hospice Transaction Form" Then
		Call navigate_to_MAXIS_screen("CASE", "NOTE")
		note_row = 5                                'beginning of listed case notes
		one_year_ago = DateAdd("yyyy", -1, date)    'we will look back 1 year
		Do
			EMReadScreen note_date, 8, note_row, 6      'reading the date
			EMReadScreen note_title, 55, note_row, 25   'reading the header
			note_title = trim(note_title)

			If left(note_title, 41) = "*** HOSPICE TRANSACTION FORM RECEIVED ***" Then      'if the note is for a Hospice form
				EmWriteScreen "X", note_row, 3      'open the note
				transmit

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
		Call navigate_to_MAXIS_screen ("STAT", "MEMB")      'Going to MEMB for M01 to see if there is a date of death - as that would be the exit date
		EMReadScreen date_of_death, 10, 19, 42
		date_of_death = replace(date_of_death, " ", "/")
		If IsDate(date_of_death) = TRUE Then hosp_exit_date = date_of_death
	End If
Next


'Future Iteration -Capturing count of each form so we can iterate the necessary form dialogs -This works well to have after all of the form selection dialogs. Then it doesn't count weird in the do/loop.
' For form_added = 0 to Ubound(form_type_array, 2)
' 	If form_type_array(form_type_const, form_added) = "Asset Statement" Then asset_count = asset_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Authorization to Release Information (ATR)" Then atr_count = atr_count + 1
' 	If form_type_array(form_type_const, form_added) = "AREP (Authorized Rep)" Then arep_count = arep_count + 1
' 	If form_type_array(form_type_const, form_added) = "Change Report Form" Then change_count = change_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Employment Verification Form (EVF)" Then evf_count = evf_count + 1  
' 	If form_type_array(form_type_const, form_added) = "Hospice Transaction Form" Then hosp_count = hosp_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Interim Assistance Agreement (IAA)" Then iaa_count = iaa_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Interim Assistance Authorization- SSI" Then iaa_ssi_count = iaa_ssi_count + 1 
' 	If form_type_array(form_type_const, form_added) = "LTC-1503" Then ltc_1503_count = ltc_1503_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Medical Opinion Form (MOF)" Then mof_count = mof_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Minnesota Transition Application Form (MTAF)" Then mtaf_count = mtaf_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Professional Statement of Need (PSN)" Then psn_count = psn_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Residence and Shelter Expenses Release Form" Then sf_count = sf_count + 1 
' 	If form_type_array(form_type_const, form_added) = "Special Diet Information Request (MFIP and MSA)" Then diet_count = diet_count + 1 
' Next
' MsgBox "checking count of each form" & vbcr & "Asset count " & asset_count & vbcr & "ATR count " & atr_count & vbcr & "AREP " & arep_count & vbcr & "chng " & change_count & vbcr & "evf " & evf_count & vbcr & "hosp " & hosp_count & vbcr & "iaa " & iaa_count & vbcr & "iaa-ssi " & iaa_ssi_count & vbcr & "ltc-1503 " & ltc_1503_count & vbcr & "mof " & mof_count & vbcr & "mtaf " & mtaf_count & vbcr & "psn " & psn_count & vbcr & "sf " & sf_count & vbcr & "diet " & diet_count	'TEST


'DIALOG DISPLAYING FORM SPECIFIC INFORMATION===========================================================================
'Displays individual dialogs for each form selected via checkbox or dropdown. Do/Loops allows us to jump around/are more flexible than For/Next 
form_count = 0
Do	
	Do
		Do
			err_msg = ""
			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 456, 300, "Documents Received"
				If form_type_array(form_type_const, form_count) = "Asset Statement" then Call asset_dialog
				If form_type_array(form_type_const, form_count) = "Authorization to Release Information (ATR)" Then Call atr_dialog
				If form_type_array(form_type_const, form_count) = "AREP (Authorized Rep)" then Call arep_dialog
				If form_type_array(form_type_const, form_count) = "Change Report Form" Then Call change_dialog
				If form_type_array(form_type_const, form_count) = "Employment Verification Form (EVF)" Then Call evf_dialog
				If form_type_array(form_type_const, form_count) = "Hospice Transaction Form" Then Call hospice_dialog
				If form_type_array(form_type_const, form_count) = "Interim Assistance Agreement (IAA)" Then Call iaa_dialog
				If form_type_array(form_type_const, form_count) = "Interim Assistance Authorization- SSI" Then Call iaa_ssi_dialog
				If form_type_array(form_type_const, form_count) = "LTC-1503" Then Call ltc_1503_dialog
				If form_type_array(form_type_const, form_count) = "Medical Opinion Form (MOF)" Then Call mof_dialog
				If form_type_array(form_type_const, form_count) = "Minnesota Transition Application Form (MTAF)" Then Call mtaf_dialog
				If form_type_array(form_type_const, form_count) = "Professional Statement of Need (PSN)" Then Call psn_dialog
				If form_type_array(form_type_const, form_count) = "Residence and Shelter Expenses Release Form" Then Call sf_dialog
				If form_type_array(form_type_const, form_count) = "Special Diet Information Request (MFIP and MSA)" Then Call diet_dialog
				
				btn_pos = 45		'variable to iterate down for each necessary button
				''Future Iteration - handle to uniquely identify multiples of the same form by adding count to the button name
				For current_form = 0 to Ubound(form_type_array, 2) 		'This iterates through the array and creates buttons for each form selected from top down. Also stores button name and number in the array based on form name selected. 
					If form_type_array(form_type_const, current_form) = "Asset Statement" then
						form_type_array(btn_name_const, form_count) = "ASSET"
						form_type_array(btn_number_const, form_count) = 400
						PushButton 395, btn_pos, 45, 15, "ASSET", asset_btn
						'PushButton 395, btn_pos, 45, 15, "ASSET-" & asset_count, asset_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
						' MsgBox "asset name" & form_type_array(form_type_const, current_form) 'TEST
						' MsgBox "asset btn" & form_type_array(btn_name_const, form_count)	'TEST
						' MsgBox "asset numb" & form_type_array(btn_number_const, form_count) 'TEST
					End If
					If form_type_array(form_type_const, current_form) = "Authorization to Release Information (ATR)" Then 
						form_type_array(btn_name_const, form_count) = "ATR"
						form_type_array(btn_number_const, form_count) = 401
						PushButton 395, btn_pos, 45, 15, "ATR", atr_btn
						'PushButton 395, btn_pos, 45, 15, "ATR-" & atr_count, atr_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					' 	MsgBox "atr name" & form_type_array(form_type_const, current_form) 'TEST
					' 	MsgBox "atr btn" & form_type_array(btn_name_const, form_count)	'TEST
					' 	MsgBox "atr numb" & form_type_array(btn_number_const, form_count) 'TEST
					End If
					If form_type_array(form_type_const, current_form) = "AREP (Authorized Rep)" then 
						form_type_array(btn_name_const, form_count) = "AREP"
						form_type_array(btn_number_const, form_count) = 402
						PushButton 395, btn_pos, 45, 15, "AREP", arep_btn
						'PushButton 395, btn_pos, 45, 15, "AREP-" & arep_count, arep_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Change Report Form"  then 
						form_type_array(btn_name_const, form_count) = "CHNG"
						form_type_array(btn_number_const, form_count) = 403
						PushButton 395, btn_pos, 45, 15, "CHNG", change_btn 
						'PushButton 395, btn_pos, 45, 15, "CHNG-" & change_count, change_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Employment Verification Form (EVF)"  then 
						form_type_array(btn_name_const, form_count) = "EVF"
						form_type_array(btn_number_const, form_count) = 404		
						PushButton 395, btn_pos, 45, 15, "EVF", evf_btn 
						'PushButton 395, btn_pos, 45, 15, "EVF-" & evf_count, evf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Hospice Transaction Form"  then 
						form_type_array(btn_name_const, form_count) = "HOSP"
						form_type_array(btn_number_const, form_count) = 405
						PushButton 395, btn_pos, 45, 15, "HOSP", hospice_btn 
						'PushButton 395, btn_pos, 45, 15, "HOSP-" & hosp_count, hospice_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Interim Assistance Agreement (IAA)"  then 
						form_type_array(btn_name_const, form_count) = "IAA"
						form_type_array(btn_number_const, form_count) = 406
						PushButton 395, btn_pos, 45, 15, "IAA", iaa_btn
						'PushButton 395, btn_pos, 45, 15, "IAA-" & iaa_count, iaa_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Interim Assistance Authorization- SSI" then 
						form_type_array(btn_name_const, form_count) = "IAA-SSI"
						form_type_array(btn_number_const, form_count) = 407
						PushButton 395, btn_pos, 45, 15, "IAA-SSI", iaa_ssi_btn 
						'PushButton 395, btn_pos, 45, 15, "IAA-SSI-" & iaa_ssi_count, iaa_ssi_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "LTC-1503" then 
						form_type_array(btn_name_const, form_count) = "LTC-1503"
						form_type_array(btn_number_const, form_count) = 408
						PushButton 395, btn_pos, 45, 15, "LTC-1503", ltc_1503_btn 
						'PushButton 395, btn_pos, 45, 15, "LTC-1503-" & ltc_1503_count, ltc_1503_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Medical Opinion Form (MOF)" then 
						form_type_array(btn_name_const, form_count) = "MOF"
						form_type_array(btn_number_const, form_count) = 409
						PushButton 395, btn_pos, 45, 15, "MOF", mof_btn 
						'PushButton 395, btn_pos, 45, 15, "MOF-" & mof_count, mof_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Minnesota Transition Application Form (MTAF)" then 
						form_type_array(btn_name_const, form_count) = "MTAF"
						form_type_array(btn_number_const, form_count) = 410
						PushButton 395, btn_pos, 45, 15, "MTAF", mtaf_btn
						'PushButton 395, btn_pos, 45, 15, "MTAF-" & mtaf_count, mtaf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Professional Statement of Need (PSN)" then 
						form_type_array(btn_name_const, form_count) = "PSN"
						form_type_array(btn_number_const, form_count) = 411
						PushButton 395, btn_pos, 45, 15, "PSN", psn_btn 
						'PushButton 395, btn_pos, 45, 15, "PSN-" & psn_count, psn_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Residence and Shelter Expenses Release Form" then 
						form_type_array(btn_name_const, form_count) = "SF"
						form_type_array(btn_number_const, form_count) = 412
						PushButton 395, btn_pos, 45, 15, "SF", sf_btn
						'PushButton 395, btn_pos, 45, 15, "SF-" & sf_count, sf_btn 'TEST - example of adding number to name of button
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = "Special Diet Information Request (MFIP and MSA)" then 
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

				
			EndDialog
			dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
			cancel_confirmation
			
			'TODO: error handling 
			' Special diet: 
				'If denied, state reason for ineligibility and date benefits are no longer issued in Comments field or create an additional field
				'Buttons	
				' If IsDate(diet_effective_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Effective Date."
				' If IsDate(diet_date_received) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				' If diet_member_number = "Select" Then err_msg = err_msg & vbNewLine & "* Select the resident for special diet."
				' If diet_diagnosis = "" Then err_msg = err_msg & vbNewLine & "* Enter diagnosis"
				' If diet_1_dropdown <>"" and diet_relationship_1_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 1 relationship"
				' If diet_2_dropdown <>"" and diet_relationship_2_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 2 relationship"
				' If diet_3_dropdown <>"" and diet_relationship_3_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 3 relationship"
				' If diet_4_dropdown <>"" and diet_relationship_4_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 4 relationship"

				' If diet_relationship_1_dropdown <>"" and diet_1_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 1 diet"
				' If diet_relationship_2_dropdown <>"" and diet_2_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 2 diet"
				' If diet_relationship_3_dropdown <>"" and diet_3_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 3 diet"
				' If diet_relationship_4_dropdown <>"" and diet_4_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet 4 diet"
				' If diet_length_diet = "" Then err_msg = err_msg & vbNewLine & "* Enter length of prescribed diet"
				' If diet_status_dropdown = "" Then err_msg = err_msg & vbNewLine & "* Select Diet Status"
				' If ButtonPressed = diet_link_CM_special_diet Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_002312"
				' If ButtonPressed = diet_SP_referrals Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Processing_Special_Diet_Referral.aspx"
		

			'Hospice 
				' If ButtonPressed = hosp_TE0207081_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/sites/hs-es-poli-temp/Documents%202/Forms/AllItems.aspx?id=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%202%2FTE%2002%2E07%2E081%20HOSPICE%20CASES%2Epdf&parent=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%202"
				' If ButtonPressed = hosp_SP_hospice_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Hospice.aspx"
				' If IsDate(hosp_effective_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Effective Date." 
				' If IsDate(hosp_date_received) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Document Date." 
				' If hosp_resident_name = "Select" Then err_msg = err_msg & vbNewLine & "* Select the resident that is in hospice."
				' If trim(hosp_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the Hospice the client entered."       'hospice name required
				' If IsDate(hosp_entry_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Hospice Entry."   'entry date also required
				' If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
						
			'IAA-SSI
				'If iaa_ssi_member_dropdown = "Select" Then err_msg = err_msg & vbNewLine & "* Select the resident from the dropdown."
				'If iaa_ssi_type_assistance = "" Then err_msg = err_msg & vbNewLine & "* Select type of interim assistance."
				'TODO: handling for checkboxes iaa_ssi_within_30_checkbox, iaa_ssi_outside_30_checkbox,
				'If ButtonPressed = iaa_ssi_CM121203_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00121203"
				'If ButtonPressed = iaa_ssi_sp_btn Then "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/STAT_PBEN.aspx"
			

			' IAA
				'If iaa_member_dropdown = "Select" Then err_msg = err_msg & vbNewLine & "* Select the resident from the dropdown."
				'If iaa_type_assistance = "" Then err_msg = err_msg & vbNewLine & "* Select type of interim assistance."
				'TODO: handling for checkboxes iaa_within_30_checkbox, iaa_outside_30_checkbox,
				'If ButtonPressed = iaa_sp_btn Then "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/STAT_PBEN.aspx"
			'ATR 
				'If IsDate(atr_effective_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Effective Date."
				' If IsDate(atr_date_received) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				' If atr_member_dropdown = "Select" Then err_msg = err_msg & vbNewLine & "* Select a member from the Member dropdown."
				'If IsDate(atr_start_date) = FALSE Then Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the Start Date."
				'If IsDate(atr_end_date) = FALSE Then Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the End Date."
				'If atr_authorization_type = "" Then err_msg = err_msg & vbNewLine & "* Select a valid authorization type from the dropdown"
				'If atr_contact_type = "" Then err_msg = err_msg & vbNewLine & "* Select a valid contact type from the dropdown"
				'If atr_name = "" Then err_msg = err_msg & vbNewLine & "* Enter contact name"
				'If atr_address = "" Then err_msg = err_msg & vbNewLine & "* Enter address"
				'If atr_city = "" Then err_msg = err_msg & vbNewLine & "* Enter city"
				'If atr_state = "" Then err_msg = err_msg & vbNewLine & "* Select a state"
				'If atr_zipcode = "" Then err_msg = err_msg & vbNewLine & "* Enter zip code"
				'If atr_phone_number = "" Then err_msg = err_msg & vbNewLine & "* Enter phone number"
				'If atr_eval_treat_checkbox and atr_coor_serv_checkbox and atr_elig_serv_checkbox and atr_court_checkbox and atr_other_checkbox and atr_other = "" Then err_msg = err_msg & vbNewLine & "* At least one checkbox must be checked indicating the use of the requested records"
				'If atr_other_checkbox = checked and atr_comments <> "" Then err_msg = err_msg & vbNewLine & "* Other checkbox was checked. You are required to specify details in the box below."
			' change TODO 


			Call dialog_movement	'function to move throughout the dialogs
						
						'MsgBox "i" & i  TEST
						'MsgBox "form type-form count @ end" & form_type_array(form_type_const, form_count) 'TEST
							
		Loop until err_msg = ""
	'Loop until form_count > Ubound(form_type_array, 2)
	Loop until ButtonPressed = complete_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
'MsgBox "Date Effective: " & chng_effective_date + vbCr + "Date Received" & chng_date_received + vbCr + "Address" & chng_address_notes + vbCr + "Household Members" & chng_household_notes + vbCr + "Assets" & chng_asset_notes + vbCr + "Vehicles" & chng_vehicles_notes + vbCr + "Income" & chng_income_notes + vbCr + "Shelter" & chng_shelter_notes + vbCr + "Other" & chng_other_change_notes + vbCr + "Action Taken" & chng_actions_taken + vbCr + "Other Notes" & chng_other_notes + vbCr + "Verifs Requested" & chng_verifs_requested + vbCr + "The changes client reports" & chng_changes_continue		'TEST

'CASE NOTE===========================================================================
'TODO- Hospice: Must keep the same header otherwise reading of past case notes won't work/continue -explore how to create separate case notes for each form
		'Call write_variable_in_CASE_NOTE("*** HOSPICE TRANSACTION FORM RECEIVED ***")
		
'Asset Statement Case Notes
If form_type_array(form_type_const, form_count) = "Asset Statement" then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** ASSET STATEMENT RECEIVED ***")
End If
' 'ATR Case Notes
If form_type_array(form_type_const, form_count) = "Authorization to Release Information (ATR)" Then 
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
End If

'AREP Case Notes
If form_type_array(form_type_const, form_count) = "AREP (Authorized Rep)" then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** AREP Received ***")
End If
'Change Reported Case Note
If form_type_array(form_type_const, form_count) = "Change Report Form" Then 
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
End If

'EVF Case Notes
If form_type_array(form_type_const, form_count) = "Employment Verification Form (EVF)" Then 
	Call start_a_blank_case_note
	Call write_variable_in_case_note("*** EVF FORM RECEIVED ***")
End If
'Hospice Case Notes
If form_type_array(form_type_const, form_count) = "Hospice Transaction Form" Then 
	Call start_a_blank_case_note
	Call write_variable_in_case_note("*** HOSPICE TRANSACTION FORM RECEIVED ***")
	Call write_bullet_and_variable_in_CASE_NOTE("Effective Date", hosp_effective_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Client", hosp_resident_name)
	Call write_bullet_and_variable_in_CASE_NOTE("Hospice Name", hosp_name)
	Call write_bullet_and_variable_in_CASE_NOTE("NPI Number", hosp_npi_number)
	Call write_bullet_and_variable_in_CASE_NOTE("Date of Entry", hosp_entry_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Exit Date", hosp_exit_date)
	'Call write_bullet_and_variable_in_MMIS_NOTE("Exit due to", exit_cause)         'This field is not currently in use so commented out - workers are testing, may add it back in
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS updated as of", hosp_mmis_updated_date)
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS not updated due to", hosp_reason_not_updated)
	Call write_bullet_and_variable_in_CASE_NOTE("Notes", hosp_other_notes)
End If
'IAA Case Notes
If form_type_array(form_type_const, form_count) = "Interim Assistance Agreement (IAA)" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** INTERIM ASSISTANCE AGREEMENT RECEIVED ***")
	CALL write_bullet_and_variable_in_case_note("Effective Date", iaa_effective_date)
	CALL write_bullet_and_variable_in_case_note("Date Received", iaa_date_received)
	CALL write_bullet_and_variable_in_case_note("Household Member", iaa_member_dropdown)
	'If iaa_within_30_checkbox = checked Then CALL write_variable_in_case_note("Signed within 30 days of receiving Combined Application Form or Change Report Form.")	'TODO FIX
	'If iaa_outside_30_checkbox = checked Then CALL write_variable_in_case_note("NOT signed within 30 days of receiving Combined Application Form or Change Report Form.")	'TODO FIX
	CALL write_bullet_and_variable_in_case_note("Other benefits resident may be eligible for", "   " & iaa_benefits_1 & "   " & iaa_benefits_2 & "   " & iaa_benefits_3 & "   " & iaa_benefits_4)
	CALL write_bullet_and_variable_in_case_note("Notes", iaa_comments)
End If

'IAA-SSI Case Notes
If form_type_array(form_type_const, form_count) = "Interim Assistance Authorization- SSI" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** INTERIM ASSISTANCE AGREEMENT-SSI RECEIVED ***")
	CALL write_bullet_and_variable_in_case_note("Effective Date", iaa_ssi_effective_date)
	CALL write_bullet_and_variable_in_case_note("Date Received", iaa_ssi_date_received)
	CALL write_bullet_and_variable_in_case_note("Household Member", iaa_ssi_member_dropdown)
	CALL write_bullet_and_variable_in_case_note("Assistance Type", iaa_ssi_type_assistance)
	'If iaa_ssi_within_30_checkbox = checked Then CALL write_variable_in_case_note("Signed within 30 days of receiving Combined Application Form or Change Report Form.")	'TODO FIX
	'If iaa_ssi_outside_30_checkbox = checked Then CALL write_variable_in_case_note("NOT signed within 30 days of receiving Combined Application Form or Change Report Form.")	'TODO FIX
	CALL write_bullet_and_variable_in_case_note("Notes", iaa_ssi_comments)
End If

'LTC 1503 Case Notes
If form_type_array(form_type_const, form_count) = "LTC-1503" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** LTC-1503 FORM RECEIVED ***")
End IF

'MOF Case Notes
If form_type_array(form_type_const, form_count) = "Medical Opinion Form (MOF)" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** MEDICAL OPINION FORM RECEIVED ***")
End If

'MTAF Case Notes
If form_type_array(form_type_const, form_count) = "Minnesota Transition Application Form (MTAF)" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** MINNESOTA TRANSITION APPLICATION RECEIVED ***")
End If

'PSN Case Notes
If form_type_array(form_type_const, form_count) = "Professional Statement of Need (PSN)" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** PROFESSIONAL STATEMENT OF NEED RECEIVED ***")
End If

'SF Case Notes
If form_type_array(form_type_const, form_count) = "Residence and Shelter Expenses Release Form" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** SHELTER FORM RECEIVED ***")
End If

'Special Diet Case Notes
If form_type_array(form_type_const, form_count) = "Special Diet Information Request (MFIP and MSA)" Then 
	Call start_a_blank_case_note
	CALL write_variable_in_case_note("*** SPECIAL DIET FORM RECEIVED ***")	
	CALL write_bullet_and_variable_in_case_note("Date Effective", diet_effective_date)					
	CALL write_bullet_and_variable_in_case_note("Date Received", diet_date_received)					
	CALL write_bullet_and_variable_in_case_note("Member", diet_member_number)							'required
	CALL write_bullet_and_variable_in_case_note("Diagnosis", diet_diagnosis)
	CALL write_bullet_and_variable_in_case_note("  Diet 1", diet_1_dropdown & "- " & diet_relationship_1_dropdown)	'required	'TODO: figure out why this populates on every casenote
	CALL write_bullet_and_variable_in_case_note("  Diet 2", diet_2_dropdown & "- " & diet_relationship_2_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("  Diet 3", diet_3_dropdown & "- " & diet_relationship_3_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("  Diet 4", diet_4_dropdown & "- " & diet_relationship_4_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("Last exam date", diet_date_last_exam)
	CALL write_bullet_and_variable_in_case_note("Diet Length", diet_length_diet)							'required
	CALL write_bullet_and_variable_in_case_note("Person following treatment plan", diet_treatment_plan_dropdown)
	If diet_status_dropdown = "Incomplete" then
		CALL write_bullet_and_variable_in_case_note("Diet approved/denied", diet_status_dropdown & "- form returned to client")
	Else
		CALL write_bullet_and_variable_in_case_note("Diet approved/denied", diet_status_dropdown)
	End If 
	CALL write_bullet_and_variable_in_case_note("Prognosis", diet_prognosis)
	CALL write_bullet_and_variable_in_case_note("Comments",diet_comments)
	CALL write_variable_in_case_note("   ")
	
	CALL write_variable_in_case_note("---")
	CALL write_variable_in_case_note(worker_signature)
End If


'TODO- look at what script this was for
'If we checked to TIKL out, it goes to TIKL and sends a TIKL
' IF tikl_nav_check = 1 THEN
' 	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
' 	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
' 	EMSetCursor 9, 3
' END IF

script_end_procedure ("Success! The script has ended. ")



'ADDING A NEW FORM TO SCRIPT TO DO LIST 
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