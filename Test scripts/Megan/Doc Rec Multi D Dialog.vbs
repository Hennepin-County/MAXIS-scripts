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
review_selections 	= 203
clear_button		= 204
next_btn			= 205
previous_btn		= 206

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
		Text 60, 25, 45, 10, MAXIS_case_number
		EditBox 175, 20, 45, 15, atr_effective_date
		EditBox 310, 20, 45, 15, atr_date_received
		EditBox 30, 65, 270, 15, atr_Q1
		EditBox 30, 85, 270, 15, atr_Q2
		EditBox 30, 105, 270, 15, atr_Q3
		EditBox 30, 125, 270, 15, atr_Q4
		Text 5, 5, 220, 10, "AUTHORIZATION TO RELEASE INFORMATION (ATR)"
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
	Text 60, 20, 45, 10, "MAXIS_case_number"
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
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, hosp_effective_date
			EditBox 310, 20, 45, 15, hosp_date_received
			EditBox 30, 65, 270, 15, hosp_Q1
			EditBox 30, 85, 270, 15, hosp_Q2
			EditBox 30, 105, 270, 15, hosp_Q3
			EditBox 30, 125, 270, 15, hosp_Q4			
			Text 5, 5, 220, 10, "Hospice Transaction Form"
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

function iaa_dialog()
		Text 60, 25, 45, 10, MAXIS_case_number
		EditBox 175, 20, 45, 15, iaa_effective_date
		EditBox 310, 20, 45, 15, iaa_date_received
		EditBox 30, 65, 270, 15, iaa_Q1
		EditBox 30, 85, 270, 15, iaa_Q2
		EditBox 30, 105, 270, 15, iaa_Q3
		EditBox 30, 125, 270, 15, iaa_Q4		
		Text 5, 5, 220, 10, "Interim Assistance Agreement (IAA)"
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

function iaa_ssi_dialog()
		Text 60, 25, 45, 10, MAXIS_case_number
			EditBox 175, 20, 45, 15, iaa_ssi_effective_date
			EditBox 310, 20, 45, 15, iaa_ssi_date_received
			EditBox 30, 65, 270, 15, iaa_ssi_Q1
			EditBox 30, 85, 270, 15, iaa_ssi_Q2
			EditBox 30, 105, 270, 15, iaa_ssi_Q3
			EditBox 30, 125, 270, 15, iaa_ssi_Q4			
			Text 5, 5, 220, 10, "Interim Assistance Authorization- SSI"
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
	DropListBox 55, 35, 65, 15, "", diet_member_number 'TODO: Need to populate member number here
	EditBox 55, 55, 290, 15, diet_diagnosis
	DropListBox 55, 85, 110, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_1_dropdown
	DropListBox 185, 85, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_1_dropdown
	DropListBox 55, 100, 110, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_2_dropdown
	DropListBox 185, 100, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_2_dropdown
	DropListBox 55, 115, 110, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_3_dropdown
	DropListBox 185, 115, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_3_dropdown
	DropListBox 55, 130, 110, 20, ""+chr(9)+"Anti-dumping"+chr(9)+"Controlled protein 40-60 grams"+chr(9)+"Controlled protein <40 grams"+chr(9)+"Gluten free"+chr(9)+"High Protein"+chr(9)+"High residue"+chr(9)+"Hypoglycemic"+chr(9)+"Ketogenic"+chr(9)+"Lactose free"+chr(9)+"Low cholesterol"+chr(9)+"Pregnancy/Lactation", diet_4_dropdown
	DropListBox 185, 130, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive", diet_relationship_4_dropdown
	EditBox 75, 160, 55, 15, diet_date_last_exam
	DropListBox 130, 180, 35, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_treatment_plan_dropdown
	EditBox 270, 180, 55, 15, diet_length_diet
	DropListBox 105, 200, 60, 15, ""+chr(9)+"Approved"+chr(9)+"Denied", diet_approved_denied_dropdown
	EditBox 50, 220, 290, 15, diet_prognosis
	EditBox 50, 240, 290, 15, diet_comments
	Text 395, 35, 45, 10, "    --Forms--"
	Text 5, 5, 220, 10, "SPECIAL DIET INFORMATION REQUEST (MFIP and MSA)"
	Text 5, 20, 50, 10, "Case Number:"
	Text 60, 20, 45, 10, "MAXIS_case_number"
	Text 125, 20, 50, 10, "Effective Date:"
	Text 245, 20, 60, 10, "Document Date:"
	Text 20, 40, 35, 10, "Member"
	Text 20, 60, 35, 10, "Diagnosis"
	Text 55, 75, 85, 10, "Select Applicable Diet"
	Text 185, 75, 95, 10, "Relationship between diets"
	Text 30, 85, 20, 10, "Diet 1"
	Text 30, 100, 20, 10, "Diet 2"
	Text 30, 115, 20, 10, "Diet 3"
	Text 30, 130, 20, 10, "Diet 4"
	Text 15, 165, 60, 10, "Date of last exam"
	Text 15, 185, 115, 10, "Is person following treament plan?"
	Text 185, 185, 85, 10, "Length of Prescribed Diet"
	Text 15, 205, 85, 10, "Diet approved or denied?"
	Text 15, 225, 35, 10, "Prognosis"
	Text 15, 245, 35, 10, "Comments"
end function

Dim diet_effective_date, diet_date_received, diet_member_number, diet_diagnosis, diet_1_dropdown, diet_2_dropdown, diet_3_dropdown, diet_4_dropdown, diet_relationship_1_dropdown, diet_relationship_2_dropdown, diet_relationship_3_dropdown, diet_relationship_4_dropdown, diet_date_last_exam, diet_treatment_plan_dropdown, diet_approved_denied_dropdown, diet_length_diet, diet_prognosis, diet_comments	'Special Diet Variables


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
				
			If ButtonPressed = add_button and form_type <> "" Then				'If statement to know when to store the information in the array
				ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
				form_type_array(form_type_const, form_count) = Form_type		'Storing form name in the array		
				form_count= form_count + 1 										'incrementing in the array
			End If
				
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
				MsgBox "form type" & form_type 'TEST
				MsgBox "Form selections cleared." & vbNewLine & "-Make new selection."	'Notify end user that entries were cleared.
			End If
			
			If form_count = 0 and ButtonPressed = Ok Then err_msg = "-Add forms to process or select cancel to exit script"		'If form_count = 0, then no forms have been added to doc rec to be processed.	
			'TODO - Limit to one of each form max 
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
						PushButton 70, 180, 70, 15, "Review Selections", review_selections
						CancelButton 150, 180, 40, 15
					Text 5, 5, 200, 10, "Select documents received, then Review Selections."
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
								
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and iaa_ssi_checkbox = unchecked and ltc_1503_checkbox = unchecked and mof_checkbox = unchecked and mtaf_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""	
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If		
Loop Until ButtonPressed = Ok




' 'Capturing count of each form so we can iterate the necessary form dialogs
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
	' Do
	' 	Do
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
				'TODO: handle to uniquely identify multiples of the same form by adding count to the button name
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
				
				PushButton 395, 255, 50, 15, "Previous", previous_btn ' Previous button to navigate from one form to the previous one. TODO: Determine if we need  more handling around this. 
				PushButton 395, 275, 50, 15, "Next Form", next_btn	'Next button to navigate from one form to the next. TODO: Determine if we need more handling around this. 
				EndDialog
				dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation
				
'TODO: error handling 
	
		Call dialog_movement	'function to move throughout the dialogs
	

'MsgBox "i" & i  TEST
'MsgBox "form type-form count @ end" & form_type_array(form_type_const, form_count) 'TEST
Loop until form_count > Ubound(form_type_array, 2)

'MsgBox "Date Effective: " & chng_effective_date + vbCr + "Date Received" & chng_date_received + vbCr + "Address" & chng_address_notes + vbCr + "Household Members" & chng_household_notes + vbCr + "Assets" & chng_asset_notes + vbCr + "Vehicles" & chng_vehicles_notes + vbCr + "Income" & chng_income_notes + vbCr + "Shelter" & chng_shelter_notes + vbCr + "Other" & chng_other_change_notes + vbCr + "Action Taken" & chng_actions_taken + vbCr + "Other Notes" & chng_other_notes + vbCr + "Verifs Requested" & chng_verifs_requested + vbCr + "The changes client reports" & chng_changes_continue		'TEST


'TODO: Case Notes for each dialog
'CASE NOTE===========================================================================

Call start_a_blank_case_note	'Navigates to case note
CALL write_variable_in_case_note ("--Docs Received--")		'TODO: determine what else to add to the header



'Casenote template
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()
	' CALL write_bullet_and_variable_in_case_note()

'Asset Statement Case Notes
'ATR Case Notes
'AREP Case Notes
'Change Reported Case Note
	CALL write_bullet_and_variable_in_case_note("CHANGE REPORTED--Date Effective", chng_effective_date)
	CALL write_bullet_and_variable_in_case_note("  Notable changes reported", chng_notable_change)
	CALL write_bullet_and_variable_in_case_note("  Date Received", chng_date_received)
	CALL write_bullet_and_variable_in_case_note("  Address", chng_address_notes)
	CALL write_bullet_and_variable_in_case_note("  Household Members", chng_household_notes)
	CALL write_bullet_and_variable_in_case_note("  Assets", chng_asset_notes)
	CALL write_bullet_and_variable_in_case_note("  Vehicles", chng_vehicles_notes)
	CALL write_bullet_and_variable_in_case_note("  Income", chng_income_notes)
	CALL write_bullet_and_variable_in_case_note("  Shelter", chng_shelter_notes)
	CALL write_bullet_and_variable_in_case_note("  Other", chng_other_change_notes)
	CALL write_bullet_and_variable_in_case_note("  Action Taken", chng_actions_taken)
	CALL write_bullet_and_variable_in_case_note("  Other Notes", chng_other_notes)
	CALL write_bullet_and_variable_in_case_note("  Verifs Requested", chng_verifs_requested)
	CALL write_bullet_and_variable_in_case_note("  The changes client reports", chng_changes_continue)
	CALL write_variable_in_case_note("   ")
'EVF Case Notes
'Hospice Case Notes
'IAA Case Notes
'IAA-SSI Case Notes
'LTC 1503 Case Notes
'MOF Case Notes
'MTAF Case Notes
'PSN Case Notes
'SF Case Notes
'Special Diet Case Notes
	CALL write_bullet_and_variable_in_case_note("SPECIAL DIET--Date Effective", diet_effective_date)	
	CALL write_bullet_and_variable_in_case_note("  Date Received", diet_date_received)					
	CALL write_bullet_and_variable_in_case_note("  Member", diet_member_number)							'required
	CALL write_bullet_and_variable_in_case_note("  Diagnosis", diet_diagnosis)
	CALL write_bullet_and_variable_in_case_note("    Diet 1", diet_1_dropdown & diet_relationship_1_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("    Diet 2", diet_2_dropdown & diet_relationship_2_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("    Diet 3", diet_3_dropdown & diet_relationship_3_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("    Diet 4", diet_4_dropdown & diet_relationship_4_dropdown)	'required
	CALL write_bullet_and_variable_in_case_note("  Last exam date", diet_date_last_exam)
	CALL write_bullet_and_variable_in_case_note("  Diet Length", diet_length_diet)							'required
	CALL write_bullet_and_variable_in_case_note("  Person following treatment plan", diet_treatment_plan_dropdown)
	CALL write_bullet_and_variable_in_case_note("  Diet approved/denied", diet_approved_denied_dropdown)
	CALL write_bullet_and_variable_in_case_note("  Prognosis", diet_prognosis)
	CALL write_bullet_and_variable_in_case_note("  Comments",diet_comments)



CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to TIKL and sends a TIKL
IF tikl_nav_check = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF

script_end_procedure ("Success! The script has ended. ")



'ADDING A NEW FORM TO SCRIPT TO DO LIST 
'Define Count Var
'Define BTN Var
'Define Dialog Function 
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