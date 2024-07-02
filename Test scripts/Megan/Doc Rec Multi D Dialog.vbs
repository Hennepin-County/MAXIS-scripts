'STATS GATHERING=============================================================================================================
name_of_script = "NOTES - DOCUMENTS RECEIVED-TESTING.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the an actual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation applicable to your script.
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/03/2024", "Restructured the dialog, added additional forms and allows user to specific document date for each form. Users select specific documents received. Forms include: Assets, ATR, AREP, Change, Hospice, IAA, LTC1503, MOF, MTAF, PSN, Shelter, Special Diet, Other", "Megan Geissler, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("11/14/2022", "Added a review of PROG for an interview date for the MTAF option selection.", "Casey Love, Hennepin County")
call changelog_update("03/16/2022", "Removed Interview Date field and added a link to supports for the MFIP Orientation.", "Casey Love, Hennepin County")
call changelog_update("03/03/2022", "Removed DVD Orientation option in the MTAF form supports.", "Ilse Ferris")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
Call changelog_update("01/03/2020", "Added new functionality to ask about accepting documents in ECF as a reminder at the end of the script.", "Casey Love, Hennepin County")
Call changelog_update("09/25/2019", "Bug Fix - script would error/stop if case was stuck in background. Added a number of checks to be sure case is not in background so the script run can continue.", "Casey Love, Hennepin County")
Call changelog_update("07/29/2019", "Bug fix - script was not identifying document information as complete when only SHEL editbox was filled.", "Casey Love, Hennepin County")
Call changelog_update("07/27/2019", "Functionality for specific forms:  Assets, MOF, AREP, LTC 1503, and MTAF. Form functionality can be accessed by checkboxes on the main dialog though all document detail can still be added in theeditboxes on the main dialog.", "Casey Love, Hennepin County")
call changelog_update("03/08/2019", "EVF received functionality added. This used to be a seperate script and will now be a part of documents received.", "Casey Love, Hennepin County")
call changelog_update("01/03/2017", "Added HSR scanner option for Hennepin County users only.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect "" 'Connects to BlueZone
Call check_for_MAXIS(false)


'TODO LIST - ADDING A NEW FORM TO SCRIPT=========================================================================== 
'Define Form Name
'Define Count Var
'Define BTN Var
'Define Dialog Function 
'Dim Variables
'Dialog Movement- BTN and Form Name
'Err Handling
'Drop Down Selection Dialog
'Checkbox VAR
'Checkbox Dialog
'Array Form Capture
'Capture Count
'Call Dialog and define current_dialog
'Define form name, btn number,and btn name, add error handling to dialog
'Define docs_rec and end_msg 
'TIKLs
'Case Note 

'DEFINING CONSTANTS, ARRAY and BUTTONS===========================================================================
'Define Constants
const form_type_const   = 0
const btn_name_const    = 1
const btn_number_const	= 2
const count_of_form		= 3
const the_last_const	= 4

'Defining form array capturing form names, button names, button numbers
Dim form_type_array()		'Defining 1D array
ReDim form_type_array(the_last_const, 0)	'Redefining array so we can resize it 
form_count = 0				'Counter for array should start with 0
all_form_array = "*"
false_count = 0 

'Define Form Names
asset_form_name 	= "Asset Statement"
atr_form_name		= "Authorization to Release Information (ATR)"
arep_form_name		= "AREP (Authorized Rep)"
change_form_name	= "Change Report Form"
evf_form_name		= "Employment Verification Form (EVF)"
hosp_form_name		= "Hospice Transaction Form"
iaa_form_name		= "Interim Assistance Agreement (IAA and IAA-SSI)"
ltc_1503_form_name	= "LTC-1503"
mof_form_name		= "Medical Opinion Form (MOF)"
mtaf_form_name		= "MN Transition Application Form (MTAF)"
psn_form_name		= "Professional Statement of Need (PSN)"
sf_form_name		= "Proof of Shelter/Residence Expenses"
diet_form_name		= "Special Diet Information Request"
other_form_name		= "**Other form/form not listed**"

'Buttons Defined
add_button 			= 201
all_forms 			= 202
clear_button		= 204
next_btn			= 205
previous_btn		= 206
complete_btn		= 207
none_btn			= 208

asset_btn			= 400
atr_btn				= 401
arep_btn			= 402
change_btn 			= 403
evf_btn				= 404
hospice_btn			= 405
iaa_btn				= 406
ltc_1503_btn		= 408
mof_btn				= 409
mtaf_btn			= 410
psn_btn				= 411
sf_btn				= 412
diet_btn			= 413
other_btn			= 414


'Define resource buttons
iaa_CM121203_btn				= 2000	
iaa_te021214_btn				= 2002
iaa_smi_btn 					= 2003
diet_link_CM_special_diet		= 2004
diet_SP_referrals				= 2005
hosp_TE0207081_btn				= 2006
hosp_SP_hospice_btn				= 2007
psn_CM1315_btn					= 2008
psn_TE1817_btn					= 2009
psn_hss_btn						= 2010
psn_mhm_btn						= 2011
psn_hsss_btn					= 2012
mtaf_cm101801_btn				= 2013
mtaf_cm0510_btn					= 2014
mtaf_mfip_orientation_info_btn	= 2015
mtaf_cm15121206_btn				= 2016
diet_link_CM_special_diet		= 2017
msg_show_instructions_btn 		= 2018


'ASSET CODE-START
Dim ASSETS_ARRAY()
ReDim ASSETS_ARRAY(update_panel, 0)

Const ast_panel         = 0
Const ast_owner         = 1
Const ast_ref_nbr       = 2
Const ast_instance      = 3
Const ast_type          = 4
Const ast_balance       = 5
Const ast_verif         = 6
Const ast_number        = 7
Const ast_wthdr_YN      = 8
Const ast_wdrw_penlty   = 9
Const ast_wthdr_verif   = 10
Const ast_jnt_owner_YN  = 11
Const ast_own_ratio      = 12
Const ast_othr_ownr_one = 13
Const ast_othr_ownr_two = 14
Const ast_othr_ownr_thr = 15
Const ast_owner_signed  = 16
Const apply_to_CASH     = 17
Const apply_to_SNAP     = 18
Const apply_to_HC       = 19
Const apply_to_GRH      = 20
Const apply_to_IVE      = 21
Const ast_location      = 22
Const ast_model         = 23
Const ast_make          = 24
Const ast_year          = 25
Const ast_trd_in        = 26
Const ast_loan_value    = 27
Const ast_value_srce    = 28
Const ast_amt_owed      = 29
Const ast_owe_verif     = 30
Const ast_owed_date     = 31
Const ast_hc_benefit    = 32
Const ast_bal_date      = 33
Const ast_verif_date    = 34
Const ast_next_inrst_date = 35
Const ast_owe_YN        = 36
Const ast_use           = 37
Const update_date       = 38
Const cnote_panel       = 39
Const ast_csv           = 40
Const ast_face_value    = 41
Const ast_share_note    = 42
Const ast_note          = 43

Const update_panel      = 44

Dim client_list_array
'ASSET CODE-END

'ADDR/HEST/SHEL START 
Dim ALL_SHEL_PANELS_ARRAY()
ReDim ALL_SHEL_PANELS_ARRAY(shel_entered_notes_const, 0)

const shel_ref_number_const 		= 00
const shel_exists_const 			= 01
const memb_btn_const				= 02
const hud_sub_yn_const 				= 03
const shared_yn_const 				= 04
const paid_to_const 				= 05
const rent_retro_amt_const 			= 06
const rent_retro_verif_const 		= 07
const rent_prosp_amt_const 			= 08
const rent_prosp_verif_const 		= 09
const lot_rent_retro_amt_const 		= 10
const lot_rent_retro_verif_const 	= 11
const lot_rent_prosp_amt_const 		= 12
const lot_rent_prosp_verif_const	= 13
const mortgage_retro_amt_const 		= 14
const mortgage_retro_verif_const 	= 15
const mortgage_prosp_amt_const 		= 16
const mortgage_prosp_verif_const 	= 17
const insurance_retro_amt_const 	= 18
const insurance_retro_verif_const 	= 19
const insurance_prosp_amt_const 	= 20
const insurance_prosp_verif_const 	= 21
const tax_retro_amt_const 			= 22
const tax_retro_verif_const 		= 23
const tax_prosp_amt_const 			= 24
const tax_prosp_verif_const 		= 25
const room_retro_amt_const 			= 26
const room_retro_verif_const 		= 27
const room_prosp_amt_const 			= 28
const room_prosp_verif_const 		= 29
const garage_retro_amt_const 		= 30
const garage_retro_verif_const 		= 31
const garage_prosp_amt_const 		= 32
const garage_prosp_verif_const 		= 33
const subsidy_retro_amt_const 		= 34
const subsidy_retro_verif_const 	= 35
const subsidy_prosp_amt_const 		= 36
const subsidy_prosp_verif_const 	= 37
const attempted_update_const 		= 38
const person_shel_checkbox 			= 39
const person_shel_button			= 40
const person_age_const 				= 41
const original_panel_info_const		= 42
const new_shel_pers_total_amt_const = 43
const new_shel_pers_total_amt_type_const = 44
const shel_entered_notes_const		= 45

ADDR_page_btn					= 3018
SHEL_page_btn					= 3019
HEST_page_btn					= 3020

ADDR_dlg_page 					= 3021
SHEL_dlg_page 					= 3022
HEST_dlg_page 					= 3023

update_information_btn 			= 500
save_information_btn			= 501
clear_mail_addr_btn				= 502
clear_phone_one_btn				= 503
clear_phone_two_btn				= 504
clear_phone_three_btn			= 505
clear_all_btn					= 506
view_total_shel_btn				= 507
update_household_percent_button = 508
housing_change_continue_btn 	= 509
housing_change_overview_btn 	= 510
housing_change_addr_update_btn 	= 511
housing_change_shel_update_btn	= 512
housing_change_shel_details_btn = 513

enter_shel_one_btn 		= 550
enter_shel_two_btn		= 551
enter_shel_three_btn 	= 552

update_addr 			= False
update_shel 			= False
update_hest 			= False
caf_answer_droplist 	= ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
show_totals 			= True

total_current_rent 			= 0
total_current_taxes 		= 0
total_current_lot_rent 		= 0
total_current_room 			= 0
total_current_mortgage 		= 0
total_current_garage 		= 0
total_current_insurance		= 0
total_current_subsidy 		= 0
total_paid_to 				= ""
total_paid_by_household 	= 100
total_paid_by_others 		= 0
'END ADDR/SHEL/HEST


'FUNCTIONS DEFINED===========================================================================
'ASSET CODE-START
function update_ACCT_panel_from_dialog()
    EMWriteScreen "                    ", 7, 44
    EMWriteScreen "                    ", 8, 44
    EMWriteScreen "        ", 10, 46
    EMWriteScreen "  ", 11, 44
    EMWriteScreen "  ", 11, 47
    EMWriteScreen "  ", 11, 50
    EMWriteScreen "        ", 12, 46

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 2), 6, 44
    EMWriteScreen ASSETS_ARRAY(ast_number, asset_counter), 7, 44
    EMWriteScreen ASSETS_ARRAY(ast_location, asset_counter), 8, 44
    EMWriteScreen ASSETS_ARRAY(ast_balance, asset_counter), 10, 46
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 10, 64
    Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_bal_date, asset_counter), 0, 11, 44)
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_YN, asset_counter), 12, 64
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_verif, asset_counter), 12, 72
    EMWriteScreen ASSETS_ARRAY(ast_wdrw_penlty, asset_counter), 12, 46
    EMWriteScreen ASSETS_ARRAY(apply_to_CASH, asset_counter), 14, 50
    EMWriteScreen ASSETS_ARRAY(apply_to_SNAP, asset_counter), 14, 57
    EMWriteScreen ASSETS_ARRAY(apply_to_HC, asset_counter), 14, 64
    EMWriteScreen ASSETS_ARRAY(apply_to_GRH, asset_counter), 14, 72
    EMWriteScreen ASSETS_ARRAY(apply_to_IVE, asset_counter), 14, 80
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 15, 44
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 15, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 15, 80
    If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) <> "" Then
        EMWriteScreen left(ASSETS_ARRAY(ast_next_inrst_date, asset_counter), 2), 17, 57
        EMWriteScreen right(ASSETS_ARRAY(ast_next_inrst_date, asset_counter), 2), 17, 60
    Else
        EMWriteScreen "  ", 17, 57
        EMWriteScreen "  ", 17, 60
    End If
end function

function update_SECU_panel_from_dialog()
    EMWriteScreen "            ", 7, 50
    EMWriteScreen "                    ", 8, 50
    EMWriteScreen "        ", 10, 52
    EMWriteScreen "  ", 11, 35
    EMWriteScreen "  ", 11, 38
    EMWriteScreen "  ", 11, 41
    EMWriteScreen "        ", 12, 52
    EMWriteScreen "        ", 13, 52

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 2), 6, 50
    EMWriteScreen ASSETS_ARRAY(ast_number, asset_counter), 7, 50
    EMWriteScreen ASSETS_ARRAY(ast_location, asset_counter), 8, 50
    EMWriteScreen ASSETS_ARRAY(ast_csv, asset_counter), 10, 52
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 11, 50
    Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_bal_date, asset_counter), 0, 11, 35)
    EMWriteScreen ASSETS_ARRAY(ast_face_value, asset_counter), 12, 52
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_YN, asset_counter), 13, 72
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_verif, asset_counter), 13, 80
    EMWriteScreen ASSETS_ARRAY(ast_wdrw_penlty, asset_counter), 13, 52
    EMWriteScreen ASSETS_ARRAY(apply_to_CASH, asset_counter), 15, 50
    EMWriteScreen ASSETS_ARRAY(apply_to_SNAP, asset_counter), 15, 57
    EMWriteScreen ASSETS_ARRAY(apply_to_HC, asset_counter), 15, 64
    EMWriteScreen ASSETS_ARRAY(apply_to_GRH, asset_counter), 15, 72
    EMWriteScreen ASSETS_ARRAY(apply_to_IVE, asset_counter), 15, 80
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 16, 44
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 80
end function

function update_CARS_panel_from_dialog()
    EMWriteScreen "                ", 8, 43
    EMWriteScreen "                ", 8, 66
    EMWriteScreen "         ", 9, 45
    EMWriteScreen "         ", 9, 62
    EMWriteScreen "         ", 12, 45
    EMWriteScreen "  ", 13, 43
    EMWriteScreen "  ", 13, 46
    EMWriteScreen "  ", 13, 49

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 1), 6, 43
    EMWriteScreen ASSETS_ARRAY(ast_year, asset_counter), 8, 31
    EMWriteScreen ASSETS_ARRAY(ast_make, asset_counter), 8, 43
    EMWriteScreen ASSETS_ARRAY(ast_model, asset_counter), 8, 66
    EMWriteScreen ASSETS_ARRAY(ast_trd_in, asset_counter), 9, 45
    EMWriteScreen ASSETS_ARRAY(ast_loan_value, asset_counter), 9, 62
    EMWriteScreen left(ASSETS_ARRAY(ast_value_srce, asset_counter), 1), 9, 80
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 10, 60
    EMWriteScreen ASSETS_ARRAY(ast_amt_owed, asset_counter), 12, 45
    EMWriteScreen left(ASSETS_ARRAY(ast_owe_verif, asset_counter), 1), 12, 60
    If ASSETS_ARRAY(ast_owed_date, asset_counter) <> "" Then Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_owed_date, asset_counter), 0, 13, 43)
    EMWriteScreen left(ASSETS_ARRAY(ast_use, asset_counter), 1), 15, 43
    EMWriteScreen ASSETS_ARRAY(ast_hc_benefit, asset_counter), 15, 76
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 16, 43
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 80
end function

function cancel_continue_confirmation(skip_functionality)

    skip_functionality = FALSE
    If ButtonPressed = 0 then       'this is the cancel button
        cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no information provided so far will be updated or noted. If you choose 'No' the update for THIS FORM will be cancelled and rest of the script will continue." & vbNewLine & vbNewLine & "YES - Stop the script entirely." & vbNewLine & "NO - Do not stop the script entrirely, just cancel the entry of this form information."& vbNewLine & "CANCEL - I didn't mean to cancel at all. (Cancel my cancel)", vbQuestion + vbYesNoCancel, "Clarify Cancel")
        If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")     'ends the script entirely
        If cancel_clarify = vbNo Then skip_functionality = TRUE
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
    End if

end function

function asset_dialog_DHS6054_and_update_asset_panels()	' DHS6054 captures additional info on assets. Update Asset panel: Read ACCT, SECU, CARS based on dialog selection and allows user to update said panels- this is MF specific.
	If asset_err_msg = "" Then
		If asset_dhs_6054_checkbox = checked Then 
			BeginDialog Dialog1, 0, 0, 336, 235, "DHS 6054 Form Details"
				EditBox 20, 40, 310, 15, box_one_info
				EditBox 20, 85, 310, 15, box_two_info
				EditBox 20, 125, 310, 15, box_three_info
				ComboBox 20, 170, 80, 15, client_dropdown_CB, signed_by_one
				EditBox 115, 170, 50, 15, signed_one_date
				ComboBox 20, 190, 80, 15, client_dropdown_CB, signed_by_two
				EditBox 115, 190, 50, 15, signed_two_date
				ComboBox 20, 210, 80, 15, client_dropdown_CB, signed_by_three
				EditBox 115, 210, 50, 15, signed_three_date
				ButtonGroup ButtonPressed
					OkButton 280, 185, 50, 15
					CancelButton 280, 205, 50, 15
				Text 15, 115, 145, 10, "Information provided about Vehicles:"
				Text 120, 155, 35, 10, "On (date):"
				Text 15, 25, 280, 10, "Information provided about Bank Accounts, Debit Accounts, or Certificates of Deposit:"
				Text 5, 5, 330, 10, "Assets for SNAP/Cash are self attested and are reported on this form (DHS 6054)"
				Text 20, 155, 40, 10, "Signed By:"
				Text 15, 70, 280, 10, "Information provided aboutStocks, Bonds, Pensions, or Retirement Accounts:"
			EndDialog

			Do
				Do
					err_msg = ""
					dialog Dialog1
					Call cancel_confirmation
					If signed_by_one <> "Select or Type" and signed_one_date = "" Then err_msg = err_msg & vbNewLine & "Date required for signature one"
					If signed_by_two <> "Select or Type" and signed_two_date = "" Then err_msg = err_msg & vbNewLine & "Date required for signature two"
					If signed_by_three <> "Select or Type" and signed_three_date = "" Then err_msg = err_msg & vbNewLine & "Date required for signature three"
					If Err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
				Loop until err_msg = ""
				Call check_for_password(are_we_passworded_out)
			Loop until are_we_passworded_out = FALSE
		End If
		
		highest_asset = asset_counter	'TODO: Why is this necessary?
		If asset_update_panels_checkbox = checked Then
			'end_msg = end_msg & vbNewLine & "Asset detail entered."
			MAXIS_footer_month = CM_mo
			MAXIS_footer_year = CM_yr

			Do
				Call back_to_SELF
				Call MAXIS_background_check

				found_the_panel = FALSE
				panel_found = FALSE
				update_panel_type = "NONE - I'm all done"
				snap_is_yes = FALSE
				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
				'Dialog to chose the panel type'
				BeginDialog Dialog1, 0, 0, 176, 85, "Type of panel to update"
				DropListBox 15, 25, 155, 45, "NONE - I'm all done"+chr(9)+"Existing ACCT"+chr(9)+"New ACCT"+chr(9)+"Existing SECU"+chr(9)+"New SECU"+chr(9)+"Existing CARS"+chr(9)+"New CARS", update_panel_type
				EditBox 90, 45, 20, 15, MAXIS_footer_month
				EditBox 115, 45, 20, 15, MAXIS_footer_year
				ButtonGroup ButtonPressed
					OkButton 120, 65, 50, 15
				Text 10, 10, 125, 10, "What panel would you like to update?"
				Text 15, 50, 65, 10, "Footer Month/Year"
				EndDialog

				Do
					Do
						err_msg = ""
						dialog Dialog1
						cancel_confirmation
						If update_panel_type = "Existing ACCT" AND acct_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known ACCT panels, cannot update an 'Existing ACCT' panel."
						If update_panel_type = "Existing SECU" AND secu_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known SECU panels, cannot update an 'Existing SECU' panel."
						If update_panel_type = "Existing CARS" AND cars_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known CARS panels, cannot update an 'Existing CARS' panel."
						If Err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
					Loop until err_msg = ""
					Call check_for_password(are_we_passworded_out)
				Loop until are_we_passworded_out = FALSE

				panel_type = right(update_panel_type, 4)
				skip_this_panel = FALSE
				If panel_type = "ACCT" Then
					If update_panel_type = "Existing ACCT" Then
						Do
							Call navigate_to_MAXIS_screen("STAT", "ACCT")
							EMReadScreen navigate_check, 4, 2, 44
							EMWaitReady 0, 0
						Loop until navigate_check = "ACCT"
						For each member in HH_member_array
							Call write_value_and_transmit(member, 20, 76)

							EMReadScreen acct_versions, 1, 2, 78
							If acct_versions <> "0" Then
								EMWriteScreen "01", 20, 79
								transmit
								Do
									is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")

									If is_this_the_panel = vbYes Then found_the_panel = TRUE

									If found_the_panel = TRUE then
										current_member = member
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
									'EMReadScreen acct_panel_not_exist, 14, 24, 13
								Loop until reached_last_ACCT_panel = "ENTER A VALID" 'OR acct_panel_not_exist = "DOES NOT EXIST"
							End If
							If found_the_panel = TRUE then Exit For
						Next

						EMReadScreen current_instance, 1, 2, 73
						current_instance = "0" & current_instance
						For the_asset = 0 to UBound(ASSETS_ARRAY, 2)
							'MsgBox "the asset" & the_asset &  "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
							If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
								asset_counter = the_asset
								If ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y" Then count_cash_checkbox = checked
								If ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then count_snap_checkbox = checked
								If ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y" Then count_hc_checkbox = checked
								If ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y" Then count_grh_checkbox = checked
								If ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y" Then count_ive_checkbox = checked
								share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								Exit For
							End If
						Next
					ElseIf update_panel_type = "New ACCT" Then
						ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
					End If

					If share_ratio_num = "" Then share_ratio_num = "1"
					If share_ratio_denom = "" Then share_ratio_denom = "1"
					If asset_dhs_6054_checkbox = checked AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"

					ASSETS_ARRAY(ast_verif_date, asset_counter) = asset_date_received
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					'Dialog to fill the ACCT panel
					BeginDialog Dialog1, 0, 0, 271, 235, "New ACCT panel for Case #" & MAXIS_case_number
					DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
					DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+ACCT_type_list, ASSETS_ARRAY(ast_type, asset_counter)
					EditBox 75, 50, 105, 15, ASSETS_ARRAY(ast_number, asset_counter)
					EditBox 75, 70, 105, 15, ASSETS_ARRAY(ast_location, asset_counter)
					EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_balance, asset_counter)
					EditBox 160, 90, 50, 15, ASSETS_ARRAY(ast_bal_date, asset_counter)
					DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_verif, asset_counter)
					If asset_dhs_6054_checkbox = unchecked Then EditBox 75, 130, 50, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
					CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
					CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
					CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
					CheckBox 230, 70, 30, 10, "GRH", count_grh_checkbox
					CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
					EditBox 75, 165, 50, 15, ASSETS_ARRAY(ast_wdrw_penlty, asset_counter)
					DropListBox 75, 185, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_wthdr_verif, asset_counter)
					EditBox 215, 125, 15, 15, share_ratio_num
					EditBox 240, 125, 15, 15, share_ratio_denom
					ComboBox 170, 160, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_one, asset_counter)
					ComboBox 170, 175, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_two, asset_counter)
					ComboBox 170, 190, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
					EditBox 75, 210, 50, 15, ASSETS_ARRAY(ast_next_inrst_date, asset_counter)
					ButtonGroup ButtonPressed
						OkButton 160, 215, 50, 15
						CancelButton 215, 215, 50, 15
					Text 10, 15, 60, 10, "Owner of Account:"
					Text 20, 35, 50, 10, "Account Type:"
					Text 15, 55, 60, 10, "Account Number:"
					Text 10, 75, 60, 10, "Account Location:"
					Text 40, 95, 30, 10, "Balance:"
					Text 130, 95, 25, 10, "As of:"
					Text 30, 115, 40, 10, "Verification:"
					GroupBox 225, 10, 40, 90, "Count:"
					GroupBox 20, 150, 140, 55, "Withdrawal Penalty"
					Text 40, 170, 30, 10, "Amount:"
					Text 30, 190, 40, 10, "Verification:"
					If asset_dhs_6054_checkbox = unchecked Then Text 35, 135, 35, 10, "Verif Date:"
					GroupBox 165, 110, 100, 100, "Additional Owner(s)"
					Text 170, 130, 40, 10, "Share Ratio:"
					Text 170, 145, 50, 10, "Other owners:"
					Text 5, 215, 65, 10, "Next Interest Date:"
					Text 235, 125, 5, 10, "/"
					EndDialog

					Do
						Do
							err_msg = ""
							dialog Dialog1
							Call cancel_continue_confirmation(skip_this_panel)
							ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(ASSETS_ARRAY(ast_wdrw_penlty, asset_counter))
							ASSETS_ARRAY(ast_number, asset_counter) = trim(ASSETS_ARRAY(ast_number, asset_counter))
							ASSETS_ARRAY(ast_location, asset_counter) = trim(ASSETS_ARRAY(ast_location, asset_counter))
							ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = trim(ASSETS_ARRAY(ast_next_inrst_date, asset_counter))
							share_ratio_num = trim(share_ratio_num)
							share_ratio_denom = trim(share_ratio_denom)
							If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the bank account. The person must be listed in the household to have a new ACCT panel added."
							If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of account this is."
							If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Select the verification source for this account."
							If ASSETS_ARRAY(ast_number, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_number, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The account number is too long."
							If ASSETS_ARRAY(ast_location, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_location, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The location name is too long."
							If IsNumeric(ASSETS_ARRAY(ast_balance, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance should be entered as a number."
							If IsDate(ASSETS_ARRAY(ast_bal_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance effective date should be entered as a date."
							If IsNumeric(share_ratio_num) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_num > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
							End If
							If IsNumeric(share_ratio_denom) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_denom > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
							End If
							If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_next_inrst_date, asset_counter)) <> 5 Then err_msg = err_msg & vbNewLine & "* The next interest date should be entered in the format MM/YY."

							If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
							Else
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
								If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
							End If
							If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
							If skip_this_panel = TRUE Then
								err_msg = ""
								If update_panel_type = "New ACCT" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
							End If
							If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
						Loop until err_msg = ""
						Call check_for_password(are_we_passworded_out)
					Loop until are_we_passworded_out = FALSE

					If skip_this_panel = FALSE Then
						ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)
						If count_cash_checkbox = checked Then ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y"
						If count_snap_checkbox = checked Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						If count_hc_checkbox = checked Then ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y"
						If count_grh_checkbox = checked Then ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y"
						If count_ive_checkbox = checked Then ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y"
						If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
						If share_ratio_denom = "1" Then
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
						Else
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
							ASSETS_ARRAY(ast_share_note, asset_counter) = "ACCT is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
						End If
						ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
						If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = ""
						Do
							Call navigate_to_MAXIS_screen("STAT", "ACCT")
							EMReadScreen navigate_check, 4, 2, 44
							EMWaitReady 0, 0
						Loop until navigate_check = "ACCT"
						EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
						If update_panel_type = "Existing ACCT" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
						transmit
						If update_panel_type = "New ACCT" Then
							EMWriteScreen "NN", 20, 79
							transmit
						End If
						If update_panel_type = "Existing ACCT" Then PF9
						ASSETS_ARRAY(cnote_panel, asset_counter) = checked
						ASSETS_ARRAY(ast_panel, asset_counter) = "ACCT"
						Call update_ACCT_panel_from_dialog
						actions_taken =  actions_taken & "Updated ACCT " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "
						If update_panel_type = "New ACCT" Then
							EMReadScreen the_instance, 1, 2, 73
							ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
						End If
						transmit
						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								panel_found = FALSE
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44
									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")
									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""
							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If
							Call update_ACCT_panel_from_dialog
							transmit
							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If
						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								panel_found = FALSE
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44
									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")
									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""

							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If

							Call update_ACCT_panel_from_dialog
							transmit

							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If

						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								panel_found = FALSE
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44
									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")
									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""
							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If
							Call update_ACCT_panel_from_dialog
							transmit
							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If
					End If
					if update_panel_type = "New ACCT" Then asset_counter = asset_counter + 1
					if update_panel_type = "Existing ACCT" Then asset_counter = highest_asset
				ElseIf panel_type = "SECU" Then
					If update_panel_type = "Existing SECU" Then
						Do
							Call navigate_to_MAXIS_screen("STAT", "SECU")
							EMReadScreen navigate_check, 4, 2, 45
							EMWaitReady 0, 0
						Loop until navigate_check = "SECU"
						For each member in HH_member_array
							Call write_value_and_transmit(member, 20, 76)
							EMReadScreen secu_versions, 1, 2, 78
							If secu_versions <> "0" Then
								EMWriteScreen "01", 20, 79
								transmit
								Do
									is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")
									If is_this_the_panel = vbYes Then found_the_panel = TRUE
									If found_the_panel = TRUE then
										current_member = member
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_SECU_panel, 13, 24, 2
								Loop until reached_last_SECU_panel = "ENTER A VALID"
							End If
							If found_the_panel = TRUE then Exit For
						Next

						EMReadScreen current_instance, 1, 2, 73
						current_instance = "0" & current_instance
						For the_asset  = 0 to UBound(ASSETS_ARRAY, 2)
							'MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
							If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
								asset_counter = the_asset
								If ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y" Then count_cash_checkbox = checked
								If ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then count_snap_checkbox = checked
								If ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y" Then count_hc_checkbox = checked
								If ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y" Then count_grh_checkbox = checked
								If ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y" Then count_ive_checkbox = checked
								share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								Exit For
							End If
						Next

					Else update_panel_type = "New SECU"
						ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
					End If
					If share_ratio_num = "" Then share_ratio_num = "1"
					If share_ratio_denom = "" Then share_ratio_denom = "1"
					If (asset_dhs_6054_checkbox = checked AND ASSETS_ARRAY(ast_verif, asset_counter) = "") Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"
					ASSETS_ARRAY(ast_verif_date, asset_counter) = asset_date_received
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					'Dialog to fill the SECU panel
					BeginDialog Dialog1, 0, 0, 271, 235, "New SECU panel for Case #" & MAXIS_case_number
					DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
					DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+SECU_type_list, ASSETS_ARRAY(ast_type, asset_counter)
					EditBox 75, 50, 105, 15, ASSETS_ARRAY(ast_number, asset_counter)
					EditBox 75, 70, 105, 15, ASSETS_ARRAY(ast_location, asset_counter)
					EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_csv, asset_counter)
					EditBox 160, 90, 50, 15, ASSETS_ARRAY(ast_bal_date, asset_counter)
					DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Agency Form"+chr(9)+"2 - Source Doc"+chr(9)+"3 - Phone Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prov", ASSETS_ARRAY(ast_verif, asset_counter)
					If asset_dhs_6054_checkbox = unchecked Then EditBox 95, 130, 60, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
					EditBox 95, 150, 60, 15, ASSETS_ARRAY(ast_face_value, asset_counter)
					CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
					CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
					CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
					CheckBox 230, 70, 30, 10, "GRH", count_grh_checkbox
					CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
					EditBox 75, 190, 50, 15, ASSETS_ARRAY(ast_wdrw_penlty, asset_counter)
					DropListBox 75, 210, 80, 45, "Select..."+chr(9)+"1 - Agency Form"+chr(9)+"2 - Source Doc"+chr(9)+"3 - Phone Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prov", ASSETS_ARRAY(ast_wthdr_verif, asset_counter)
					EditBox 215, 125, 15, 15, share_ratio_num
					EditBox 240, 125, 15, 15, share_ratio_denom
					ComboBox 170, 160, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_one, asset_counter)
					ComboBox 170, 175, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_two, asset_counter)
					ComboBox 170, 190, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
					ButtonGroup ButtonPressed
						OkButton 160, 215, 50, 15
						CancelButton 215, 215, 50, 15
					Text 10, 15, 60, 10, "Owner of Security:"
					Text 20, 35, 50, 10, "Security Type:"
					Text 10, 55, 60, 10, "Security Number:"
					Text 15, 75, 55, 10, "Security Name:"
					Text 10, 95, 60, 10, "Cash Value (CSV):"
					Text 25, 115, 40, 10, "Verification:"
					Text 130, 95, 25, 10, "As of:"
					If asset_dhs_6054_checkbox = unchecked Then Text 50, 135, 35, 10, "Verif Date:"
					GroupBox 225, 10, 40, 90, "Count:"
					Text 10, 155, 75, 10, "Face Value of Life Ins:"
					GroupBox 20, 175, 140, 55, "Withdrawal Penalty"
					Text 40, 195, 30, 10, "Amount:"
					Text 30, 215, 40, 10, "Verification:"
					GroupBox 165, 110, 100, 100, "Additional Owner(s)"
					Text 170, 130, 40, 10, "Share Ratio:"
					Text 170, 145, 50, 10, "Other owners:"
					Text 235, 125, 5, 10, "/"
					EndDialog

					Do
						Do
							err_msg = ""
							dialog Dialog1
							Call cancel_continue_confirmation(skip_this_panel)
							ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(ASSETS_ARRAY(ast_wdrw_penlty, asset_counter))
							ASSETS_ARRAY(ast_number, asset_counter) = trim(ASSETS_ARRAY(ast_number, asset_counter))
							ASSETS_ARRAY(ast_location, asset_counter) = trim(ASSETS_ARRAY(ast_location, asset_counter))
							ASSETS_ARRAY(ast_face_value, asset_counter) = trim(ASSETS_ARRAY(ast_face_value, asset_counter))
							share_ratio_num = trim(share_ratio_num)
							share_ratio_denom = trim(share_ratio_denom)
							If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the security. The person must be listed in the household to have a new SECU panel added."
							If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of security this is."
							If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Select the verification source for this account."
							If ASSETS_ARRAY(ast_number, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_number, asset_counter)) > 12 Then err_msg = err_msg & vbNewLine & "* The account number is too long."
							If ASSETS_ARRAY(ast_location, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_location, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The location name is too long."
							If IsNumeric(ASSETS_ARRAY(ast_csv, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance should be entered as a number."
							If IsDate(ASSETS_ARRAY(ast_bal_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance effective date should be entered as a date."
							If left(ASSETS_ARRAY(ast_type, asset_counter), 2) = "LI" Then
								If ASSETS_ARRAY(ast_face_value, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* A life insurance policy requires a face value."
								If count_snap_checkbox = checked Then
									count_snap_checkbox = unchecked
								End If
							Else
								If ASSETS_ARRAY(ast_face_value, asset_counter) <> "" Then err_msg = err_msg & vbNewLine & "* A face value amount can only be entered for a Life Insurance security."
							End If
							If IsNumeric(share_ratio_num) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_num > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
							End If
							If IsNumeric(share_ratio_denom) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_denom > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
							End If
							If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
							Else
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
								If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
							End If
							If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
							If skip_this_panel = TRUE Then
								err_msg = ""
								If update_panel_type = "New SECU" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
							End If
							If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
						Loop until err_msg = ""
						Call check_for_password(are_we_passworded_out)
					Loop until are_we_passworded_out = FALSE

					If skip_this_panel = FALSE Then
						ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)
						If count_cash_checkbox = checked Then ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y"
						If count_snap_checkbox = checked Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						If count_hc_checkbox = checked Then ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y"
						If count_grh_checkbox = checked Then ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y"
						If count_ive_checkbox = checked Then ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y"
						If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
						If share_ratio_denom = "1" Then
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
						Else
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
							ASSETS_ARRAY(ast_share_note, asset_counter) = "SECU is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
						End If
						ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
						If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = ""
						Do
							Call navigate_to_MAXIS_screen("STAT", "SECU")
							EMReadScreen navigate_check, 4, 2, 45
							EMWaitReady 0, 0
						Loop until navigate_check = "SECU"
						EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
						If update_panel_type = "Existing SECU" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
						transmit
						If update_panel_type = "New SECU" Then
							EMWriteScreen "NN", 20, 79
							transmit
						End If
						If update_panel_type = "Existing SECU" Then PF9
						ASSETS_ARRAY(cnote_panel, asset_counter) = checked
						ASSETS_ARRAY(ast_panel, asset_counter) = "SECU"
						Call update_SECU_panel_from_dialog
						actions_taken =  actions_taken & "Updated SECU " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "

						If update_panel_type = "New SECU" Then
							EMReadScreen the_instance, 1, 2, 73
							ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
						End If
						transmit

						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							panel_found = FALSE
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44

									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")

									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""

							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If

							Call update_SECU_panel_from_dialog
							transmit

							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If

						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								panel_found = FALSE
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44

									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")

									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""

							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If

							Call update_SECU_panel_from_dialog
							transmit

							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If

						If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)) = True Then
							ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)
							EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2), 20, 76
							EMWriteScreen "01", 20, 79
							transmit
							EMReadScreen total_panels, 1, 2, 78
							If total_panels = "0" Then
								EMWriteScreen "NN", 20, 79
								transmit
								panel_found = TRUE
							Else
								panel_found = FALSE
								Do
									EMReadScreen this_account_type, 2, 6, 44
									EMReadScreen this_account_number, 20, 7, 44
									EMReadScreen this_account_location, 20, 8, 44
									this_account_number = replace(this_account_number, "_", "")
									this_account_location = replace(this_account_location, "_", "")

									If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
										PF9
										panel_found = TRUE
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_ACCT_panel, 13, 24, 2
								Loop until reached_last_ACCT_panel = "ENTER A VALID"
							End If
							If panel_found = FALSE Then
								EMWriteScreen "NN", 20, 79
								transmit
							End If
							panel_found = ""

							IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
								snap_is_yes = TRUE
								ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
							End If

							Call update_ACCT_panel_from_dialog
							transmit

							If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
						End If

					End If
					if update_panel_type = "New SECU" Then asset_counter = asset_counter + 1
					if update_panel_type = "Existing SECU" Then asset_counter = highest_asset
				ElseIf panel_type = "CARS" Then
					If update_panel_type = "Existing CARS" Then
						Do
							Call navigate_to_MAXIS_screen("STAT", "CARS")
							EMReadScreen navigate_check, 4, 2, 44
							EMWaitReady 0, 0
						Loop until navigate_check = "CARS"
						For each member in HH_member_array
							Call write_value_and_transmit(member, 20, 76)

							EMReadScreen cars_versions, 1, 2, 78
							If cars_versions <> "0" Then
								EMWriteScreen "01", 20, 79
								transmit
								Do
									is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")

									If is_this_the_panel = vbYes Then found_the_panel = TRUE

									If found_the_panel = TRUE then
										current_member = member
										Exit Do
									End If
									transmit
									EMReadScreen reached_last_CARS_panel, 13, 24, 2
								Loop until reached_last_CARS_panel = "ENTER A VALID"
							End If
							If found_the_panel = TRUE then Exit For
						Next

						EMReadScreen current_instance, 1, 2, 73
						current_instance = "0" & current_instance
						For the_asset  = 0 to UBound(ASSETS_ARRAY, 2)
							'MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
							If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
								asset_counter = the_asset
								share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
								Exit For
							End If
						Next

					ElseIf update_panel_type = "New CARS" Then
						ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
					End If

					If share_ratio_num = "" Then share_ratio_num = "1"
					If share_ratio_denom = "" Then share_ratio_denom = "1"
					If asset_dhs_6054_checkbox = checked AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
					ASSETS_ARRAY(ast_verif_date, asset_counter) = asset_date_received

					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					'Dialog to fill the CARS panel.
					BeginDialog Dialog1, 0, 0, 270, 255, "New CARS panel for Case # " & MAXIS_case_number
					DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
					DropListBox 75, 30, 90, 45, "Select..."+chr(9)+CARS_type_list, ASSETS_ARRAY(ast_type, asset_counter)
					EditBox 220, 30, 40, 15, ASSETS_ARRAY(ast_year, asset_counter)
					ComboBox 75, 50, 185, 45, "Type or Select"+chr(9)+"Acura"+chr(9)+"Audi"+chr(9)+"BMW"+chr(9)+"Buick"+chr(9)+"Cadillac"+chr(9)+"Chevrolet"+chr(9)+"Chrysler"+chr(9)+"Dodge"+chr(9)+"Ford"+chr(9)+"GMC"+chr(9)+"Honda"+chr(9)+"Hummer"+chr(9)+"Hyundai"+chr(9)+"Infiniti"+chr(9)+"Isuzu"+chr(9)+"Jeep"+chr(9)+"Kia"+chr(9)+"Lincoln"+chr(9)+"Mazda"+chr(9)+"Mercedes-Benz"+chr(9)+"Mercury"+chr(9)+"Mitsubishi"+chr(9)+"Nissan"+chr(9)+"Oldsmobile"+chr(9)+"Plymouth"+chr(9)+"Pontiac"+chr(9)+"Saab"+chr(9)+"Saturn"+chr(9)+"Scion"+chr(9)+"Subaru"+chr(9)+"Suzuki"+chr(9)+"Toyota"+chr(9)+"Volkswagen"+chr(9)+"Volvo", ASSETS_ARRAY(ast_make, asset_counter)
					EditBox 75, 70, 185, 15, ASSETS_ARRAY(ast_model, asset_counter)
					EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_trd_in, asset_counter)
					DropListBox 165, 90, 95, 45, "Select..."+chr(9)+"1 - NADA"+chr(9)+"2 - Appraisal Value"+chr(9)+"3 - Client Stmt"+chr(9)+"4 - Other Document", ASSETS_ARRAY(ast_value_srce, asset_counter)
					DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Title"+chr(9)+"2 - License Reg"+chr(9)+"3 - DMV"+chr(9)+"4 - Purchase Agmt"+chr(9)+"5 - Other Document"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_verif, asset_counter)
					If asset_dhs_6054_checkbox = unchecked Then EditBox 210, 110, 50, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
					DropListBox 75, 130, 60, 45, "No"+chr(9)+"Yes", ASSETS_ARRAY(ast_hc_benefit, asset_counter)
					EditBox 75, 165, 50, 15, ASSETS_ARRAY(ast_amt_owed, asset_counter)
					EditBox 75, 185, 50, 15, ASSETS_ARRAY(ast_owed_date, asset_counter)
					DropListBox 75, 210, 80, 45, "Select..."+chr(9)+"1 - Bank/Lending Inst Stmt"+chr(9)+"2 - Private Lender Stmt"+chr(9)+"3 - Other Document"+chr(9)+"4 - Pend Out State Verif"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_owe_verif, asset_counter)
					EditBox 215, 145, 15, 15, share_ratio_num
					EditBox 240, 145, 15, 15, share_ratio_denom
					ComboBox 170, 180, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_oner, asset_counter)
					ComboBox 170, 195, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_twor, asset_counter)
					ComboBox 170, 210, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
					ButtonGroup ButtonPressed
						OkButton 170, 235, 45, 15
						CancelButton 220, 235, 45, 15
					Text 10, 15, 60, 10, "Owner of Vehicle:"
					Text 20, 35, 50, 10, "Vehicle Type:"
					Text 170, 35, 45, 10, "Vehicle Year:"
					Text 20, 55, 50, 10, "Vehicle Make:"
					Text 20, 75, 50, 10, "Vehicle Model:"
					Text 20, 95, 50, 10, "Trade In Value:"
					Text 130, 95, 25, 10, "Source:"
					Text 25, 115, 40, 10, "Verification:"
					IF asset_dhs_6054_checkbox = unchecked Then Text 170, 115, 40, 10, "Verif Date:"
					Text 15, 135, 50, 10, "HC Clt Benefit:"
					GroupBox 20, 150, 140, 80, "Amount Owed on vehicle"
					Text 40, 170, 30, 10, "Amount:"
					Text 45, 190, 20, 10, "As of:"
					Text 30, 210, 40, 10, "Verification:"
					GroupBox 165, 130, 100, 100, "Additional Owner(s)"
					Text 170, 150, 40, 10, "Share Ratio:"
					Text 170, 165, 50, 10, "Other owners:"
					Text 235, 145, 5, 10, "/"
					EndDialog

					Do
						Do
							err_msg = ""
							dialog Dialog1
							Call cancel_continue_confirmation(skip_this_panel)
							ASSETS_ARRAY(ast_year, asset_counter) = trim(ASSETS_ARRAY(ast_year, asset_counter))
							ASSETS_ARRAY(ast_make, asset_counter) = trim(ASSETS_ARRAY(ast_make, asset_counter))
							ASSETS_ARRAY(ast_model, asset_counter) = trim(ASSETS_ARRAY(ast_model, asset_counter))
							ASSETS_ARRAY(ast_trd_in, asset_counter) = trim(ASSETS_ARRAY(ast_trd_in, asset_counter))
							share_ratio_num = trim(share_ratio_num)
							share_ratio_denom = trim(share_ratio_denom)
							If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the vehicle. The person must be listed in the household to have a new SECU panel added."
							If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of vehicle this is."
							If ASSETS_ARRAY(ast_year, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the year of the vehicle."
							If len(ASSETS_ARRAY(ast_year, asset_counter)) <> 4 Then err_msg = err_msg & vbNewLine & "* The year of the vehicle needs to be in the format YYYY."
							If ASSETS_ARRAY(ast_make, asset_counter) = "Type or Select" OR ASSETS_ARRAY(ast_make, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the make of the vehicle."
							If ASSETS_ARRAY(ast_model, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the model of the vehicle."
							If IsNumeric(ASSETS_ARRAY(ast_trd_in, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The trade in value needs to be entered as a number."
							If ASSETS_ARRAY(ast_value_srce, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate from where the value was determined."
							If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Enter the verification of the vehicle."

							If ASSETS_ARRAY(ast_amt_owed, asset_counter) <> "" Then
								If IsNumeric(ASSETS_ARRAY(ast_amt_owed, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The owed amount needs to be entered as a number."
								If ASSETS_ARRAY(ast_owe_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Enter the verification of the amount that owed."
								If IsDate(ASSETS_ARRAY(ast_owed_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the effective date of the owed amount."
							End If

							If IsNumeric(share_ratio_num) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_num > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
							End If
							If IsNumeric(share_ratio_denom) = FALSE Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
							ElseIf share_ratio_denom > 9 Then
								err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
							End If

							If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
							Else
								ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
								If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
							End If
							If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
							If skip_this_panel = TRUE Then
								err_msg = ""
								If update_panel_type = "New SECU" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
							End If

							If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
						Loop until err_msg = ""
						Call check_for_password(are_we_passworded_out)
					Loop until are_we_passworded_out = FALSE

					If skip_this_panel = FALSE Then
						ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)
						If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
						If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
						If ASSETS_ARRAY(ast_owe_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_owe_verif, asset_counter) = ""
						ASSETS_ARRAY(ast_loan_value, asset_counter) = .9 * ASSETS_ARRAY(ast_trd_in, asset_counter)
						ASSETS_ARRAY(ast_loan_value, asset_counter) = round(ASSETS_ARRAY(ast_loan_value, asset_counter))
						If share_ratio_denom = "1" Then
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
						Else
							ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
							ASSETS_ARRAY(ast_share_note, asset_counter) = "CARS is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
						End If
						ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
						If ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "Yes" Then ASSETS_ARRAY(ast_hc_benefit, asset_counter)  = "Y"
						If ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "No" Then ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "N"
						Do
							Call navigate_to_MAXIS_screen("STAT", "CARS")
							EMReadScreen navigate_check, 4, 2, 44
							EMWaitReady 0, 0
						Loop until navigate_check = "CARS"
						EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
						If update_panel_type = "Existing CARS" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
						transmit
						If update_panel_type = "New CARS" Then
							EMWriteScreen "NN", 20, 79
							transmit
						End If
						If update_panel_type = "Existing CARS" Then PF9

						ASSETS_ARRAY(cnote_panel, asset_counter) = checked
						ASSETS_ARRAY(ast_panel, asset_counter) = "CARS"

						Call update_CARS_panel_from_dialog

						actions_taken =  actions_taken & "Updated CARS " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "


						If update_panel_type = "New CARS" Then
							EMReadScreen the_instance, 1, 2, 73
							ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
						End If
						transmit
					End If
					if update_panel_type = "New CARS" Then asset_counter = asset_counter + 1
					if update_panel_type = "Existing CARS" Then asset_counter = highest_asset
				End If
				highest_asset = asset_counter

			Loop until panel_type = "done"
		End If
	End If
end function
Dim signed_by_one, signed_by_two, signed_by_three, signed_one_date, signed_two_date, signed_three_date, box_one_info, box_two_info, box_three_info

function asset_dialog()
	EditBox 395, 0, 45, 15, asset_date_received
	y_pos = 40
	If acct_panels > 0 Then
		Text 10, y_pos, 95, 10, "Current ACCT panel details."
		Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
		y_pos = y_pos + 10
		For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
			If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
				Text 15, y_pos, 275, 10,  "* ACCT " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset) & " - Balance: $" & ASSETS_ARRAY(ast_balance, the_asset)
				CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
				y_pos = y_pos + 10
			End If
		Next
		y_pos = y_pos + 5
	End If
	
	If secu_panels > 0 Then
		Text 10, y_pos, 95, 10, "Current SECU panel details."
		Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
		y_pos = y_pos + 10
		For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
			If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
				Text 15, y_pos, 275, 10, "* SECU " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset)
				CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
				y_pos = y_pos + 10
			End If
		Next
		y_pos = y_pos + 5
	End If

	If cars_panels > 0 Then
		Text 10, y_pos, 95, 10, "Current CARS panel details."
		Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
		y_pos = y_pos + 10
		For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
			If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
				Text 15, y_pos, 275, 10, "* CARS " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset)
				CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
				y_pos = y_pos + 10
			End If
		Next
		y_pos = y_pos + 5
	End If

	If acct_panels = 0 AND cars_panels = 0 AND secu_panels = 0 Then 
		Text 60, 70, 250 , 10, "~~~NO CURRENT ACCT, SECU or CARS PANELS~~~"
		Text 85, 85, 250 , 10, "~~~Complete the fields below~~~"
	End If

	Text 340, 5, 55, 10, "Document Date:"
	EditBox 60, 215, 295, 15, actions_taken
	CheckBox 15, 235, 345, 15, "Check here if DHS 6054 received. Assets for SNAP/Cash are self attested and are reported on this form.", asset_dhs_6054_checkbox
	CheckBox 15, 250, 345, 15, "Check here to have the script update asset panels. (ACCT, SECU, CARS).", asset_update_panels_checkbox
	Text 5, 5, 220, 10, asset_form_name
	Text 15, 220, 45, 10, "Action Taken:"
end function
Dim asset_date_received, actions_taken, asset_dhs_6054_checkbox, asset_update_panels_checkbox
'ASSET CODE-END

function atr_dialog()
	EditBox 395, 0, 45, 15, atr_date_received		
	DropListBox 50, 40, 100, 15, HH_Memb_DropDown, atr_member_dropdown
	EditBox 205, 40, 45, 15, atr_start_date
	EditBox 300, 40, 45, 15, atr_end_date
	DropListBox 80, 60, 70, 15, ""+chr(9)+"Verbal"+chr(9)+"Written", atr_authorization_type
	DropListBox 65, 95, 60, 15, ""+chr(9)+"Organization"+chr(9)+"Person", atr_contact_type
	EditBox 160, 95, 170, 15, atr_name
	EditBox 70, 120, 75, 15, atr_phone_number
	EditBox 175, 120, 80, 15, atr_fax_number
	EditBox 45, 140, 205, 15, atr_email
	CheckBox 35, 175, 170, 10, "to continue evaluation or treatment", atr_eval_treat_checkbox
	CheckBox 35, 185, 170, 10, "to coordinate services", atr_coor_serv_checkbox
	CheckBox 35, 195, 170, 10, "to determine eligibility for assistance/service", atr_elig_serv_checkbox
	CheckBox 35, 205, 170, 10, "for court proceedings", atr_court_checkbox
	CheckBox 35, 215, 80, 10, "other (specify below)", atr_other_checkbox
	EditBox 50, 225, 90, 15, atr_other
	EditBox 50, 255, 230, 15, atr_comments
	Text 5, 5, 220, 10, atr_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 15, 45, 30, 10, "Member"
	Text 170, 45, 35, 10, "Start Date"
	Text 265, 45, 30, 10, "End Date"
	Text 15, 65, 65, 10, "Authorization Type"
	GroupBox 10, 85, 340, 85, "Contact Person/Organization"
	Text 20, 100, 45, 10, "Contact Type"
	Text 140, 100, 20, 10, "Name"
	Text 20, 125, 50, 10, "Phone Number"
	Text 160, 125, 15, 10, "Fax"
	Text 20, 145, 25, 10, "Email: "
	GroupBox 10, 165, 340, 80, "Record requested will be used: "
	Text 10, 260, 35, 10, "Comments"
end function
Dim atr_date_received, atr_member_dropdown, atr_start_date, atr_end_date, atr_authorization_type, atr_contact_type, atr_name, atr_phone_number, atr_fax_number, atr_email, atr_eval_treat_checkbox, atr_coor_serv_checkbox, atr_elig_serv_checkbox, atr_court_checkbox, atr_other_checkbox, atr_other, atr_comments

function arep_dialog()
	EditBox 395, 0, 45, 15, AREP_recvd_date
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
	EditBox 135, 195, 55, 15, arep_signature_date
	CheckBox 15, 215, 75, 10, "ID on file for AREP?", AREP_ID_check
	CheckBox 15, 230, 215, 10, "TIKL to get new HC form 12 months after date form was signed?", arep_TIKL_check
	CheckBox 260, 55, 35, 10, "SNAP", arep_SNAP_AREP_checkbox
	CheckBox 260, 65, 50, 10, "Health Care", arep_HC_AREP_checkbox
  	CheckBox 260, 75, 30, 10, "Cash", arep_CASH_AREP_checkbox
	CheckBox 255, 110, 115, 10, "AREP Req - MHCP - DHS-3437", arep_dhs_3437_checkbox
	CheckBox 255, 130, 105, 10, "AREP Req - HC12729", arep_HC_12729_checkbox
	CheckBox 255, 150, 100, 10, "SNAP AREP Choice - D405", arep_D405_checkbox
	CheckBox 255, 170, 105, 10, "AREP on CAF", arep_CAF_AREP_page_checkbox
	CheckBox 255, 190, 100, 10, "AREP on any HC App", arep_HCAPP_AREP_checkbox
	CheckBox 255, 210, 75, 10, "Power of Attorney", arep_power_of_attorney_checkbox
	Text 5, 5, 220, 10, arep_form_name
	Text 340, 5, 55, 10, "Document Date:"
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
	Text 15, 200, 115, 10, "Date form was signed (MM/DD/YY)"
	GroupBox 245, 100, 125, 155, "Specific FORM Received"
	Text 270, 120, 50, 10, "(HC)"
	Text 270, 140, 60, 10, "(Cash and SNAP)"
	Text 270, 160, 75, 10, "(SNAP and EBT Card)"
	Text 270, 180, 60, 10, "(Cash and SNAP)"
	Text 270, 200, 50, 10, "(HC)"
	Text 270, 220, 60, 10, "(HC, SNAP, Cash)"
	GroupBox 245, 45, 125, 45, "Programs Authorized for:"
	Text 250, 230, 110, 20, "Checking the FORM will indicate the programs in the CASE/NOTE"
end function
Dim  arep_name, arep_street, arep_city, arep_state, arep_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, arep_forms_to_arep_checkbox, arep_mmis_mail_to_arep_checkbox, arep_update_AREP_panel_checkbox, AREP_recvd_date, AREP_ID_check, arep_TIKL_check, arep_signature_date, arep_dhs_3437_checkbox, arep_HC_12729_checkbox, arep_D405_checkbox, arep_CAF_AREP_page_checkbox, arep_HCAPP_AREP_checkbox, arep_power_of_attorney_checkbox, arep_SNAP_AREP_checkbox, arep_HC_AREP_checkbox, arep_CASH_AREP_checkbox

function change_dialog()
	EditBox 270, 0, 45, 15, chng_effective_date
	EditBox 395, 0, 45, 15, chng_date_received		
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
	DropListBox 270, 265, 105, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", chng_changes_continue
	Text 5, 5, 220, 10, change_form_name
	Text 220, 5, 50, 10, "Effective Date:"
	Text 340, 5, 55, 10, "Document Date:"
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
	Text 180, 270, 90, 10, "The changes client reports:"

end function
Dim chng_effective_date, chng_date_received, chng_address_notes, chng_household_notes, chng_asset_notes, chng_vehicles_notes, chng_income_notes, chng_shelter_notes, chng_other_change_notes, chng_actions_taken, chng_other_notes, chng_verifs_requested, chng_changes_continue, chng_notable_change 

function evf_dialog()
	EditBox 395, 0, 45, 15, evf_date_received		
	ComboBox 70, 30, 170, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
	EditBox 85, 50, 155, 15, evf_employer
 	DropListBox 85, 70, 155, 15, HH_Memb_DropDown, evf_client
	DropListBox 105, 105, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", evf_info
	EditBox 105, 125, 60, 15, evf_info_date
	EditBox 105, 145, 60, 15, evf_request_info
	CheckBox 40, 170, 130, 10, "Create TIKL for additional information", EVF_TIKL_checkbox
	EditBox 80, 195, 240, 15, evf_actions_taken
	Text 5, 5, 220, 10, evf_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 25, 35, 40, 10, "EVF Status:"
	Text 25, 55, 55, 10, "Employer name:"
	Text 25, 75, 60, 10, "Household Memb:"
	Text 25, 95, 125, 10, "Is additional information needed?"
	Text 40, 110, 60, 10, "Addt'l Info Reqstd:"
	Text 40, 130, 55, 10, "Date Requested:"
	Text 40, 150, 65, 10, "Info Requested via:"
	Text 25, 200, 50, 10, "Actions taken:"
end function 
Dim evf_date_received, EVF_status_dropdown, evf_employer, evf_client, evf_info, evf_info_date, evf_request_info, EVF_TIKL_checkbox, evf_actions_taken

function hospice_dialog()
	EditBox 395, 0, 45, 15, hosp_date_received		
	DropListBox 85, 25, 165, 15, HH_Memb_DropDown, hosp_resident_name
	EditBox 85, 45, 205, 15, hosp_name
	EditBox 60, 65, 80, 15, hosp_npi_number
	EditBox 60, 85, 50, 15, hosp_entry_date
	EditBox 60, 105, 50, 15, hosp_exit_date
	EditBox 90, 125, 50, 15, hosp_mmis_updated_date
	EditBox 185, 145, 190, 15, hosp_reason_not_updated
	EditBox 65, 165, 310, 15, hosp_other_notes
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "TE 02.07.081", hosp_TE0207081_btn
		PushButton 65, 280, 50, 15, "MA-Hospice", hosp_SP_hospice_btn
	Text 5, 5, 220, 10, hosp_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 20, 30, 55, 10, "Resident Name:"
	Text 20, 50, 60, 10, "Name of Hospice:"
	Text 20, 70, 35, 10, "NPI Numb:"
	Text 20, 90, 40, 10, "Entry Date:"
	Text 20, 110, 35, 10, "Exit Date:"
	Text 20, 130, 70, 10, "MMIS Updated as of "
	Text 20, 150, 165, 10, "If MMIS has not yet been updated, explain reason:"
	Text 20, 170, 45, 10, "Other Notes:"
end function 
Dim hosp_date_received, hosp_resident_name, hosp_name, hosp_npi_number, hosp_entry_date, hosp_exit_date, hosp_mmis_updated_date, hosp_reason_not_updated, hosp_other_notes

function iaa_dialog()
	EditBox 395, 0, 45, 15, iaa_date_received		
	DropListBox 55, 35, 140, 15, HH_Memb_DropDown, iaa_member_dropdown
	CheckBox 25, 60, 110, 10, "Check here if IAA form received", iaa_form_received_checkbox
	DropListBox 260, 55, 95, 15, ""+chr(9)+"Initial claim"+chr(9)+"Post-eligibility", iaa_type_assistance
	CheckBox 25, 85, 125, 10, "Check here if IAA-SSI form received", iaa_ssi_form_received_checkbox
	DropListBox 260, 80, 95, 15, ""+chr(9)+"General Assistance (GA)"+chr(9)+"Housing Support (HS)", iaa_ssi_type_assistance
	EditBox 65, 110, 145, 15, iaa_benefits_1
	EditBox 65, 130, 145, 15, iaa_benefits_2
	EditBox 235, 110, 145, 15, iaa_benefits_3
	EditBox 235, 130, 145, 15, iaa_benefits_4
	CheckBox 25, 155, 125, 10, "Click here to update PBEN panel", iaa_update_pben_checkbox	
	DropListBox 95, 170, 115, 15, ""+chr(9)+"01-RSDI"+chr(9)+"02-SSI"+chr(9)+"06-Child Support"+chr(9)+"07-Alimony"+chr(9)+"08-VA Disability"+chr(9)+"09-VA Pension"+chr(9)+"10-VA Dependent Educational"+chr(9)+"11-VA Dependent Other"+chr(9)+"12-Unemployment Insurance"+chr(9)+"13-Worker's Comp"+chr(9)+"14-RR Retirement"+chr(9)+"15-Other Ret"+chr(9)+"16-Military Allot"+chr(9)+"17-EITC"+chr(9)+"18-Strike Pay"+chr(9)+"19-Other"+chr(9)+"21-SMRT", iaa_benefit_type
	DropListBox 95, 190, 115, 15, ""+chr(9)+"1-Copy of Chkstb"+chr(9)+"2-Award Letters"+chr(9)+"4-Coltrl Stmt"+chr(9)+"5-Other Document"+chr(9)+"N-No Ver Prvd", iaa_verification_dropdown
	DropListBox 95, 210, 115, 15, ""+chr(9)+"A-Appealing"+chr(9)+"D-Denied"+chr(9)+"E-Eligible"+chr(9)+"P-Pending"+chr(9)+"N-Not Appl Yet"+chr(9)+"R-Refused To Accept", iaa_disposition_code_dropdown
	EditBox 315, 170, 55, 15, iaa_referral_date
	EditBox 315, 190, 55, 15, iaa_date_applied_pben
	EditBox 315, 210, 55, 15, iaa_iaa_date
	EditBox 50, 245, 315, 15, iaa_comments
	ButtonGroup ButtonPressed
		PushButton 5, 280, 50, 15, "CM12.12.03", iaa_CM121203_btn
		PushButton 60, 280, 50, 15, "TE02.12.14", iaa_te021214_btn
		PushButton 115, 280, 120, 15, "SMI- Verify Date Applied for PBEN", iaa_smi_btn
	Text 5, 5, 220, 10, iaa_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 20, 40, 30, 10, "Member"
	Text 170, 60, 90, 10, "Type of interim assistance"
	Text 165, 85, 95, 10, "AG or HS interim assistance"
	Text 40, 100, 150, 10, "Other benefits resident might be eligible for:"
	Text 50, 175, 45, 10,"Benefit Type"
	Text 55, 195, 40, 10, "Verification"
	Text 35, 215, 60, 10, "Disposition Code"
	Text 270, 175, 45, 10, "Referral Date"
	Text 240, 195, 75, 10, "Date Applied for PBEN"
	Text 285, 215, 30, 10, "IAA Date"
	Text 15, 250, 35, 10, "Comments"
end function 
Dim iaa_date_received, iaa_member_dropdown, iaa_form_received_checkbox, iaa_type_assistance, iaa_ssi_form_received_checkbox, iaa_ssi_type_assistance, iaa_benefits_1, iaa_benefits_2, iaa_benefits_3, iaa_benefits_4, iaa_update_pben_checkbox, iaa_benefit_type, iaa_referral_date, iaa_verification_dropdown, iaa_date_applied_pben, iaa_disposition_code_dropdown, iaa_iaa_date, iaa_comments

function ltc_1503_dialog()
	EditBox 395, 0, 45, 15, ltc_1503_date_received
	EditBox 65, 35, 110, 15, ltc_1503_FACI_1503
	DropListBox 235, 35, 70, 15, ""+chr(9)+"30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", ltc_1503_length_of_stay
	DropListBox 110, 55, 40, 15, ""+chr(9)+"SNF"+chr(9)+"NF"+chr(9)+"ICF-DD"+chr(9)+"RTC", ltc_1503_level_of_care
	DropListBox 210, 55, 80, 15, ""+chr(9)+"acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-DD", ltc_1503_admitted_from
	EditBox 65, 75, 90, 15, ltc_1503_hospital_admitted_from
	EditBox 215, 75, 45, 15, ltc_1503_admit_date
	EditBox 325, 75, 45, 15, ltc_1503_discharge_date
	CheckBox 20, 95, 155, 10, "If you've processed this 1503, check here.", ltc_1503_processed_1503_checkbox
	CheckBox 15, 130, 65, 10, "Updated RLVA?", ltc_1503_updated_RLVA_checkbox
	CheckBox 85, 130, 60, 10, "Updated FACI?", ltc_1503_updated_FACI_checkbox
	CheckBox 150, 130, 50, 10, "Need 3543?", ltc_1503_need_3543_checkbox
	CheckBox 210, 130, 55, 10, "Need 3531?", ltc_1503_need_3531_checkbox
	CheckBox 270, 130, 95, 10, "Need asset assessment?", ltc_1503_need_asset_assessment_checkbox
	EditBox 130, 140, 210, 15, ltc_1503_verifs_needed
	CheckBox 15, 165, 85, 10, "Sent 3050 back to LTCF", ltc_1503_sent_3050_checkbox
	CheckBox 110, 165, 70, 10, "Sent verif req? To:", ltc_1503_sent_verif_request_checkbox
	ComboBox 180, 160, 60, 15, ""+chr(9)+"client"+chr(9)+"AREP"+chr(9)+"Client & AREP", ltc_1503_sent_request_to
	CheckBox 250, 165, 120, 10, "Sent DHS-5181 to Case Manager", ltc_1503_sent_5181_checkbox
	CheckBox 15, 195, 255, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", ltc_1503_TIKL_checkbox
	CheckBox 15, 210, 155, 10, "Check here to have the script update HCMI.", ltc_1503_HCMI_update_checkbox
	CheckBox 15, 225, 150, 10, "Check here to have the script update FACI.", ltc_1503_FACI_update_checkbox
	EditBox 105, 245, 25, 15, ltc_1503_faci_footer_month
	EditBox 135, 245, 25, 15, ltc_1503_faci_footer_year
	EditBox 250, 245, 75, 15, ltc_1503_mets_case_number
	EditBox 35, 265, 330, 15, ltc_1503_notes
	Text 5, 5, 220, 10, ltc_1503_form_name
	Text 340, 5, 55, 10, "Document Date:"
	GroupBox 5, 25, 370, 95, "Facility Info"
	Text 15, 40, 50, 10, "Facility name:"
	Text 180, 40, 50, 10, "Length of stay:"
	Text 15, 60, 95, 10, "Recommended level of care:"
	Text 160, 60, 50, 10, "Admitted from:"
	Text 15, 80, 50, 10, "Hospital Name"
	Text 160, 80, 55, 10, "Admission Date:"
	Text 270, 80, 55, 10, "Discharge Date:"
	GroupBox 5, 115, 370, 75, "Actions/Proofs"
	Text 15, 145, 115, 10, "Other proofs needed (if applicable):"
	GroupBox 5, 185, 370, 55, "Script actions"
	Text 10, 250, 95, 10, "Facility Update Month/Year:"
	Text 170, 250, 75, 10, "METS Case Number:"
	Text 10, 270, 25, 10, "Notes:"
end function
Dim ltc_1503_date_received, ltc_1503_FACI_1503, ltc_1503_length_of_stay, ltc_1503_level_of_care, ltc_1503_admitted_from, ltc_1503_hospital_admitted_from, ltc_1503_admit_date, ltc_1503_discharge_date, ltc_1503_processed_1503_checkbox, ltc_1503_updated_RLVA_checkbox, ltc_1503_updated_FACI_checkbox, ltc_1503_need_3543_checkbox, ltc_1503_need_3531_checkbox, ltc_1503_need_asset_assessment_checkbox, ltc_1503_verifs_needed, ltc_1503_sent_3050_checkbox, ltc_1503_sent_verif_request_checkbox, ltc_1503_sent_request_to, ltc_1503_sent_5181_checkbox, ltc_1503_TIKL_checkbox, ltc_1503_HCMI_update_checkbox, ltc_1503_FACI_update_checkbox, ltc_1503_faci_footer_month, ltc_1503_faci_footer_year, ltc_1503_mets_case_number, ltc_1503_notes

function mof_dialog()
	EditBox 395, 0, 45, 15, mof_date_received		
	DropListBox 45, 25, 140, 15, HH_Memb_DropDown, mof_hh_memb
	CheckBox 220, 25, 85, 10, "Client signed release?", mof_clt_release_checkbox
	EditBox 75, 45, 55, 15, mof_last_exam_date
	ComboBox 80, 65, 95, 15, "Select or Type"+chr(9)+"Less than 30 Days"+chr(9)+"Between 30 - 45 Days"+chr(9)+"More than 45 Days"+chr(9)+"No End Date Listed", mof_time_condition_will_last
	EditBox 95, 85, 55, 15, mof_doctor_date
	EditBox 100, 110, 95, 15, mof_ability_to_work
	EditBox 60, 130, 220, 15, mof_other_notes
	EditBox 60, 155, 220, 15, mof_actions_taken
	CheckBox 20, 180, 215, 10, "Check here if the MOF indicates an SSA application is needed.", mof_SSA_application_indicated_checkbox
	CheckBox 20, 195, 185, 10, "Check here if DISA will be updated as needed by TTL", mof_TTL_to_update_checkbox
	CheckBox 20, 210, 190, 10, "Check here if you sent an email to TTL/FSS DataTeam.", MOF_TTL_email_checkbox
	EditBox 100, 220, 65, 15, mof_TTL_email_date
	Text 5, 5, 220, 10, mof_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 15, 30, 30, 10, "Member:"
	Text 15, 50, 60, 10, "Date of last exam: "
	Text 15, 70, 60, 10, "Condition will last:"
	Text 15, 90, 80, 10, "Date doctor signed form: "
	Text 15, 115, 85, 10, "Member's ability to work: "
	Text 15, 135, 40, 10, "Other notes: "
	Text 285, 130, 90, 40, "...........................................      Do not enter diagnosis in case notes per PQ #16506 ............................................"
	Text 15, 160, 45, 10, "Action taken: "
	Text 40, 225, 55, 10, "Date email sent:"
end function 
Dim mof_date_received, mof_hh_memb, mof_clt_release_checkbox, mof_last_exam_date, mof_time_condition_will_last, mof_doctor_date, mof_ability_to_work, mof_other_notes, mof_actions_taken, mof_SSA_application_indicated_checkbox, mof_TTL_to_update_checkbox, MOF_TTL_email_checkbox, mof_TTL_email_date

function mtaf_dialog()	
	EditBox 395, 0, 45, 15, MTAF_date
	DropListBox 60, 20, 55, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete", MTAF_status_dropdown
	EditBox 175, 20, 45, 15, MTAF_MFIP_elig_date
	CheckBox 230, 25, 170, 10, "Check if all docs rec'vd are associated with MTAF", MTAF_note_only_checkbox 
	CheckBox 15, 40, 55, 10, "MTAF signed.", mtaf_signed_checkbox
	CheckBox 90, 40, 140, 10, "MFIP Financial Orientation completed.", mtaf_mfip_financial_orientation_checkbox
	CheckBox 230, 40, 150, 10, "Client exempt from cooperation with ES.", mtaf_ES_exemption_checkbox
	EditBox 75, 60, 205, 15, mtaf_ADDR_change
	EditBox 75, 80, 205, 15, mtaf_HHcomp_change
	EditBox 75, 100, 205, 15, mtaf_asset_change
	EditBox 95, 120, 185, 15, mtaf_earned_income_change
	EditBox 100, 140, 180, 15, mtaf_unearned_income_change
	EditBox 85, 160, 195, 15, mtaf_shelter_costs_change
	EditBox 155, 180, 50, 15, mtaf_subsidized_housing
	DropListBox 305, 180, 80, 15, "Select one..."+chr(9)+"Not subsidized"+chr(9)+"Verification provided"+chr(9)+"Verification pending", mtaf_sub_housing_droplist
	EditBox 85, 200, 95, 15, mtaf_child_adult_care_costs
	EditBox 290, 200, 100, 15, mtaf_relationship_proof
	EditBox 175, 220, 160, 15, mtaf_referred_to_OMB_PBEN
	EditBox 125, 240, 210, 15, mtaf_elig_results_fiated
	EditBox 50, 260, 125, 15, mtaf_other_notes
	EditBox 235, 260, 150, 15, mtaf_verifications_needed
	ButtonGroup ButtonPressed
		PushButton 5, 280, 55, 15, "CM 10.18.01", mtaf_cm101801_btn
		PushButton 65, 280, 45, 15, "CM 05.10", mtaf_cm0510_btn
		PushButton 115, 280, 60, 15, "CM 15.12.12.06", mtaf_cm15121206_btn
		PushButton 180, 280, 110, 15, "MFIP Orientation HSR Manual", mtaf_mfip_orientation_info_btn
	Text 5, 5, 130, 10, mtaf_form_name
	Text 355, 5, 40, 10, "MTAF date:"
	Text 15, 25, 45, 10, "MTAF status:"
	Text 125, 25, 50, 10, "MFIP elig date:"
	Text 10, 65, 65, 10, "Address changes:"
	Text 10, 85, 65, 10, "HH comp changes:"
	Text 10, 105, 65, 10, "Assets changes:"
	Text 10, 125, 85, 10, "*Earned income changes:"
	Text 10, 145, 90, 10, "Unearned income changes:"
	Text 10, 165, 70, 10, "Shelter cost changes:"
	Text 10, 185, 145, 10, "Housing subsidized amount (if applicable)?"
	Text 210, 185, 90, 10, "Subsidized housing status?"
	Text 10, 205, 75, 10, "Child/adult care costs:"
	Text 195, 205, 95, 10, "Proof of relationship on file:"
	Text 10, 225, 160, 10, "Client has been referred to apply for OMB/PBEN:"
	Text 10, 245, 115, 10, "Eligibility results fiated? If so, why:"
	Text 10, 265, 40, 10, "Other notes:"
	Text 185, 265, 50, 10, "Verifs needed:"
	GroupBox 285, 60, 105, 115, "CM 10.18.01"
	Text 290, 70, 90, 35, "*STOP WORK - Verification only necessary to verify income in the month of appl/eligibility."
	Text 290, 110, 90, 60, "*SUBSIDY - Verification of housing subsidy is a mandatory verification for MFIP. STAT must be appropriately updated to ensure accurate approval of housing grant. "
end function
Dim MTAF_note_only_checkbox, MTAF_date, MTAF_status_dropdown, MTAF_MFIP_elig_date, mtaf_signed_checkbox, mtaf_mfip_financial_orientation_checkbox, mtaf_ES_exemption_checkbox, mtaf_ADDR_change, mtaf_HHcomp_change, mtaf_asset_change, mtaf_earned_income_change, mtaf_unearned_income_change, mtaf_shelter_costs_change, mtaf_subsidized_housing, mtaf_sub_housing_droplist, mtaf_child_adult_care_costs,  mtaf_relationship_proof,  mtaf_referred_to_OMB_PBEN, mtaf_elig_results_fiated, mtaf_other_notes, mtaf_verifications_needed

function psn_dialog()
	EditBox 395, 0, 45, 15, psn_date_received	
 	DropListBox 50, 15, 100, 15, HH_Memb_DropDown, psn_member_dropdown
	DropListBox 15, 45, 105, 15, ""+CHR(9)+"Yes- At least 1 selected"+chr(9)+"No- Section NOT completed", psn_section_1_dropdown
	DropListBox 15, 60, 105, 15, ""+CHR(9)+"Yes- 1 selected"+chr(9)+"No- Section NOT completed", psn_section_2_dropdown
	DropListBox 15, 75, 105, 15, ""+CHR(9)+"Yes- At least 1 selected"+chr(9)+"No- Section NOT completed", psn_section_3_dropdown
	DropListBox 15, 90, 105, 15, ""+CHR(9)+"Yes- At least 2 selected"+chr(9)+"No- Section NOT completed", psn_section_4_dropdown
	DropListBox 15, 105, 105, 15, ""+CHR(9)+"Yes- Section completed"+chr(9)+"No- Section NOT completed", psn_section_5_dropdown
	EditBox 95, 120, 120, 15, psn_cert_prof
	EditBox 250, 120, 125, 15, psn_facility	
	CheckBox 5, 150, 185, 10, "Check here to have script update WREG/DISA panels", psn_udpate_wreg_disa_checkbox
	CheckBox 210, 150, 165, 10, "Check here to set a TIKL to update form annually", psn_tikl_checkbox
	DropListBox 65, 165, 30, 15, ""+CHR(9)+"Y"+chr(9)+"N", psn_wreg_fs_pwe
	DropListBox 195, 165, 155, 15, ""+CHR(9)+"03-Unfit for Employment"+chr(9)+"04-Resp for Care of Incapacitated Person"+chr(9)+"05-Age 60 or Older"+chr(9)+"06-Under Age 16"+chr(9)+"07-Age 16-17, Living w/ Caregiver"+chr(9)+"08-Resp for Care of Child under 6"+chr(9)+"09-Empl 30 hrs/wk or Earnings of 30 hrs/wk"+chr(9)+"10-Matching Grant Participant"+chr(9)+"11-Receiving or Applied for UI"+chr(9)+"12-Enrolled in School, Training, or Higher Ed"+chr(9)+"13-Participating in CD Program"+chr(9)+"14-Receiving MFIP"+chr(9)+"20-Pending/Receiving DWP"+chr(9)+ "15-Age 16-17, NOT Living w/ Caregiver"+chr(9)+"16-50-59 Years Old"+chr(9)+"17-Receiving RCA or GA"+chr(9)+"21-Resp for Care of Child under 18"+chr(9)+"23-Pegnant", psn_wreg_work_wreg_status
	DropListBox 65, 185, 115, 15, ""+CHR(9)+"01-Work Reg Exempt"+chr(9)+"02-Under Age 18"+chr(9)+"03-Age 50 or Over"+chr(9)+"04-Caregiver of Minor Child"+chr(9)+"05-Pregnant"+chr(9)+"06-Employed Avg of 20 hrs/wk"+chr(9)+"07-Work Experience Participant"+chr(9)+"08-ther E&T Services"+chr(9)+"09-Resides in a Waivered Area"+chr(9)+"10-ABAWD Counted Month"+chr(9)+"11-2nd-3rd Month Period of Elig"+chr(9)+"12-RCA or GA Recipient"+chr(9)+"13-ABAWD Banked Months", psn_wreg_abawd_status
	DropListBox 255, 185, 130, 20, ""+CHR(9)+"04-Permanent Ill or Incap"+chr(9)+"05-Temporary Ill or Incap"+chr(9)+"06-Care of Ill or Incap Mbr"+chr(9)+"07-Requires Services In Residence"+chr(9)+"09-Mntl Ill or Dev Disabled"+chr(9)+"10-SSI/RSDI Pend"+chr(9)+"11-Appealing SSI/RSDI Denial"+chr(9)+"12-Advanced Age"+chr(9)+"13-Learning Disability"+chr(9)+"17-Protect/Court Ordered"+chr(9)+"20-Age 16 or 17 SS Approval "+chr(9)+"25-Emancipated Minor"+chr(9)+"28-Unemployable"+chr(9)+"29-Displaced Hmkr (Ft Student)"+chr(9)+"30-Minor w/ Adult Unrelated"+chr(9)+"32-ESL, Adult/HS At least half time"+chr(9)+"35-Drug/Alcohol Addiction (DAA)"+chr(9)+"99-No Elig Basis", psn_wreg_ga_elig_status
	EditBox 65, 205, 45, 15, psn_disa_begin_date
	EditBox 255, 205, 45, 15, psn_disa_end_date
	EditBox 65, 225, 45, 15, psn_disa_cert_start
	EditBox 255, 225, 45, 15, psn_disa_cert_end
	DropListBox 65, 245, 110, 15, ""+CHR(9)+"01-RSDI Only Disability"+chr(9)+"02-RSDI Only Blindness"+chr(9)+"03-SSI, SSI/RSDI Disability"+chr(9)+"04-SSI, SSI/RSDI Blindness"+chr(9)+"06-SMRT/SSA Pend"+chr(9)+"08-SMRT Certified Blindness"+chr(9)+"09-Ill/Incapacity"+chr(9)+"10-SMRT Certified Disability", psn_disa_status
	DropListBox 255, 245, 105, 15, ""+CHR(9)+"1-DHS161/Dr Stmt"+chr(9)+"2-SMRT Certified"+chr(9)+"3-Certified For RSDI or SSI"+chr(9)+"6-Other Document"+chr(9)+"7-Professional Stmt of Need"+chr(9)+"N-No Ver Prvd", psn_disa_verif
	EditBox 55, 265, 320, 15, psn_comments
	ButtonGroup ButtonPressed
		PushButton 10, 285, 40, 15, "CM 13.15", psn_CM1315_btn
		PushButton 55, 285, 35, 15, "TE18.17", psn_TE1817_btn
		PushButton 95, 285, 30, 15, "HSS", psn_hss_btn
		PushButton 130, 285, 30, 15, "MHM", psn_mhm_btn
		PushButton 165, 285, 30, 15, "HSSS", psn_hsss_btn
	Text 5, 5, 130, 10, psn_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 15, 20, 30, 10, "Member"
	GroupBox 5, 35, 375, 105, "PSN Fields"
	Text 125, 50, 105, 10, "Section 1: Housing Situation"
	Text 125, 65, 105, 10, "Section 2: Disabling Condition"
	Text 125, 80, 150, 10, "Section 3: MA Housing Stabilization Services"
	Text 125, 95, 185, 10, "Section 4: MN Housing Support Supplemental Services"
	Text 125, 110, 220, 10, "Section 5: Transition from Residential Treatment to MN HS Program"
	Text 20, 125, 72, 10, "Certified Professional:"
	Text 225, 125, 25, 10, "Facility:"
	Text 30, 170, 30, 10, "FS PWE:"
	Text 115, 170, 80, 10, "FSET Work Reg Status: "
	Text 10, 190, 55, 10, "ABAWD Status: "
	Text 190, 190, 65, 10, "GA Elig Basis Code:"
	Text 10, 210, 55, 10, "Disa Begin Date: "
	Text 205, 210, 50, 10, "Disa End Date:"
	Text 25, 230, 40, 10, "Cert Period:"
	Text 205, 230, 50, 10, "Cert End Date:"
	Text 25, 250, 40, 10, "Disa Status: "
	Text 215, 250, 40, 10, "Verification:"
	Text 15, 270, 37, 10, "Comments:"
	Text 5, 140, 390, 10, ".............................................................................................................................................................................................."
end function 
Dim  psn_date_received, psn_member_dropdown, psn_section_1_dropdown, psn_section_2_dropdown, psn_section_3_dropdown, psn_section_4_dropdown, psn_section_5_dropdown, psn_cert_prof, psn_facility, psn_udpate_wreg_disa_checkbox, psn_tikl_checkbox, psn_wreg_fs_pwe, psn_wreg_work_wreg_status, psn_wreg_abawd_status, psn_wreg_ga_elig_status, psn_disa_begin_date, psn_disa_end_date, psn_disa_cert_start, psn_disa_cert_end, psn_disa_status, psn_disa_verif, psn_comments

function sf_dialog()	
	EditBox 395, 0, 45, 15, sf_date_received
	ComboBox 75, 45, 170, 15, "Select or Type"+chr(9)+"Contract-Deed"+chr(9)+"DHS2952 Auth Release Residence/Shelter Info"+chr(9)+"Lease"+chr(9)+"Mortgage Statement"+chr(9)+"Written Statement", sf_name_of_form
	EditBox 85, 65, 120, 15, sf_tenant_name
	EditBox 260, 65, 45, 15, sf_total_rent
	EditBox 75, 85, 25, 15, sf_adults
	EditBox 175, 85, 20, 15, sf_children
	EditBox 85, 105, 220, 15, sf_street_address
	EditBox 45, 125, 85, 15, sf_city
	EditBox 170, 125, 25, 15, sf_state
	EditBox 250, 125, 45, 15, sf_zip
	CheckBox 35, 165, 85, 10, "Electricity", sf_electrcity_checkbox
	CheckBox 35, 175, 85, 10, "Air conditioning", sf_ac_checkbox
	CheckBox 35, 185, 85, 10, "Heat", sf_heat_checkbox
	CheckBox 125, 165, 85, 10, "Gas", sf_gas_checkbox
	CheckBox 125, 175, 85, 10, "Garbage removal", sf_garbage_removal_checkbox
	CheckBox 125, 185, 85, 10, "Water and sewer", sf_water_sewer_checkbox
	CheckBox 225, 165, 85, 10, "Garage/plug-in", sf_garage_plugin_checkbox
	CheckBox 225, 175, 35, 10, "None", sf_none_checkbox
	CheckBox 225, 185, 30, 10, "Other", sf_other_checkbox
	EditBox 265, 180, 85, 15, sf_other_utility
	EditBox 65, 205, 235, 15, sf_comments
	CheckBox 30, 235, 300, 10, "Select to update ADDR, SHEL, or HEST panel", sf_update_addr_shel_hest_checkbox
	CheckBox 30, 245, 130, 10, "Check here to navigate to set a TIKL", sf_tikl_nav_check
	Text 5, 5, 220, 10, sf_form_name
	Text 340, 5, 55, 10, "Document Date:"
	GroupBox 10, 35, 350, 190, "Form Information"
	Text 30, 50, 40, 10, "Form Name"
	Text 30, 70, 45, 10, "Tenant Name"
	Text 220, 70, 40, 10, "Total Rent"
	Text 30, 90, 45, 10, "Adults in Unit"
	Text 120, 90, 55, 10, "Children in Unit"
	Text 30, 110, 50, 10, "Street Address"
	Text 30, 130, 15, 10, "City"
	Text 150, 130, 20, 10, "State"
	Text 215, 130, 30, 10, "Zipcode"
	Text 30, 155, 275, 10, "Select which utilities the tenant is responsible to pay separate from the rent amount:"
	Text 30, 210, 35, 10, "Comments"
	GroupBox 10, 220, 350, 40, "Actions"
end function
Dim sf_name_of_form, sf_date_received, sf_tenant_name, sf_total_rent, sf_adults, sf_children, sf_street_address, sf_city, sf_state, sf_zip, sf_electrcity_checkbox, sf_ac_checkbox, sf_heat_checkbox, sf_gas_checkbox, sf_garbage_removal_checkbox, sf_water_sewer_checkbox, sf_garage_plugin_checkbox, sf_none_checkbox, sf_other_checkbox, sf_other_utility, sf_update_addr_shel_hest_checkbox, sf_tikl_nav_check, sf_comments

function addr_shel_hest_panel_dialog()
	If err_msg = "" Then 
		If sf_update_addr_shel_hest_checkbox = checked Then
			Do
				err_msg = ""

				BeginDialog Dialog1, 0, 0, 555, 385, "Housing Expense Detail"

				ButtonGroup ButtonPressed
					If page_to_display = ADDR_dlg_page Then
						Text 506, 12, 60, 10, "ADDR"
						Call display_ADDR_information(update_addr, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)
					End If

					If page_to_display = SHEL_dlg_page Then 
						Text 506, 27, 60, 10, "SHEL"
						Call display_SHEL_information(update_shel, show_totals, ALL_SHEL_PANELS_ARRAY, member_selection, shel_ref_number_const, shel_exists_const, hud_sub_yn_const, shared_yn_const, paid_to_const, rent_retro_amt_const, rent_retro_verif_const, rent_prosp_amt_const, rent_prosp_verif_const, lot_rent_retro_amt_const, lot_rent_retro_verif_const, lot_rent_prosp_amt_const, lot_rent_prosp_verif_const, mortgage_retro_amt_const, mortgage_retro_verif_const, mortgage_prosp_amt_const, mortgage_prosp_verif_const, insurance_retro_amt_const, insurance_retro_verif_const, insurance_prosp_amt_const, insurance_prosp_verif_const, tax_retro_amt_const, tax_retro_verif_const, tax_prosp_amt_const, tax_prosp_verif_const, room_retro_amt_const, room_retro_verif_const, room_prosp_amt_const, room_prosp_verif_const, garage_retro_amt_const, garage_retro_verif_const, garage_prosp_amt_const, garage_prosp_verif_const, subsidy_retro_amt_const, subsidy_retro_verif_const, subsidy_prosp_amt_const, subsidy_prosp_verif_const, paid_to, percent_paid_by_household, percent_paid_by_others,  total_current_rent, total_current_lot_rent, total_current_mortgage, total_current_insurance, total_current_taxes, total_current_room, total_current_garage, total_current_subsidy, update_information_btn, save_information_btn, memb_btn_const, clear_all_btn, view_total_shel_btn, update_household_percent_button)
					End If
					
					If page_to_display = HEST_dlg_page Then 
						Text 507, 42, 60, 10, "HEST"
						Call display_HEST_information(update_hest, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, notes_on_hest, update_information_btn, save_information_btn)
					End If

					If page_to_display <> ADDR_dlg_page Then PushButton 485, 10, 65, 13, "ADDR", ADDR_page_btn
					If page_to_display <> SHEL_dlg_page Then PushButton 485, 25, 65, 13, "SHEL", SHEL_page_btn
					If page_to_display <> HEST_dlg_page Then PushButton 485, 40, 65, 13, "HEST", HEST_page_btn

					OkButton 450, 365, 50, 15
					CancelButton 500, 365, 50, 15

				EndDialog


				Dialog Dialog1
				cancel_confirmation

				If page_to_display = ADDR_dlg_page Then Call navigate_ADDR_buttons(update_addr, err_msg, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, mail_street_full, mail_city, mail_state, mail_zip, phone_one, phone_two, phone_three, type_one, type_two, type_three)
				If page_to_display = SHEL_dlg_page Then Call navigate_SHEL_buttons(update_shel, show_totals, err_var, ALL_SHEL_PANELS_ARRAY, member_selection, shel_ref_number_const, shel_exists_const, hud_sub_yn_const, shared_yn_const, paid_to_const, rent_retro_amt_const, rent_retro_verif_const, rent_prosp_amt_const, rent_prosp_verif_const, lot_rent_retro_amt_const, lot_rent_retro_verif_const, lot_rent_prosp_amt_const, lot_rent_prosp_verif_const, mortgage_retro_amt_const, mortgage_retro_verif_const, mortgage_prosp_amt_const, mortgage_prosp_verif_const, insurance_retro_amt_const, insurance_retro_verif_const, insurance_prosp_amt_const, insurance_prosp_verif_const, tax_retro_amt_const, tax_retro_verif_const, tax_prosp_amt_const, tax_prosp_verif_const, room_retro_amt_const, room_retro_verif_const, room_prosp_amt_const, room_prosp_verif_const, garage_retro_amt_const, garage_retro_verif_const, garage_prosp_amt_const, garage_prosp_verif_const, subsidy_retro_amt_const, subsidy_retro_verif_const, subsidy_prosp_amt_const, subsidy_prosp_verif_const, update_information_btn, save_information_btn, memb_btn_const, attempted_update_const, clear_all_btn, view_total_shel_btn, update_household_percent_button)

				If page_to_display = HEST_dlg_page Then Call navigate_HEST_buttons(update_hest, err_msg, update_information_btn, save_information_btn, choice_date, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, date)
				If err_msg <> "" then MsgBox "Please Resolve:" & vbCr & err_msg

				If ButtonPressed = ADDR_page_btn Then page_to_display = ADDR_dlg_page
				If ButtonPressed = SHEL_page_btn Then page_to_display = SHEL_dlg_page
				If ButtonPressed = HEST_page_btn Then page_to_display = HEST_dlg_page
			Loop until ButtonPressed = -1
		
			'Creates a new variable with MAXIS_footer_month and MAXIS_footer_year concatenated into a single date starting on the 1st of the month.
			footer_month_as_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year

			'Calculates the difference between the two dates (date of admission and footer month)
			difference_between_dates = DateDiff("m", addr_eff_date, footer_month_as_date)

			'If there's a difference between the two dates, then it backs out of the case and enters a new footer month and year, and transmits.
			If difference_between_dates <> 0 THEN
				back_to_SELF
				CALL convert_date_into_MAXIS_footer_month(addr_eff_date, MAXIS_footer_month, MAXIS_footer_year)
				EMWriteScreen MAXIS_footer_month, 20, 43
				EMWriteScreen MAXIS_footer_year, 20, 46
				Transmit
			END IF
			Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
			For shel_member = 0 to UBound(ALL_SHEL_PANELS_ARRAY, 2)
				If ALL_SHEL_PANELS_ARRAY(attempted_update_const, shel_member) = True Then
					shel_updated = true
					Call access_SHEL_panel("WRITE", ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, shel_member), ALL_SHEL_PANELS_ARRAY(hud_sub_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(shared_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(paid_to_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member))
				End If
			Next

			'here we save the the current info so that we can compare it to the original and know if it changed
			hest_current_information = all_persons_paying&"|"&all_persons_paying&"|"&choice_date&"|"&actual_initial_exp&"|"&retro_heat_ac_yn&"|"&_
			retro_heat_ac_units&"|"&retro_heat_ac_amt&"|"&retro_electric_yn&"|"&retro_electric_units&"|"&retro_electric_amt&"|"&retro_phone_yn&"|"&_
			retro_phone_units&"|"&retro_phone_amt&"|"&prosp_heat_ac_yn&"|"&prosp_heat_ac_units&"|"&prosp_heat_ac_amt&"|"&prosp_electric_yn&"|"&_
			prosp_electric_units&"|"&prosp_electric_amt&"|"&prosp_phone_yn&"|"&prosp_phone_units&"|"&prosp_phone_amt&"|"&total_utility_expense

			hest_current_information = UCASE(hest_current_information)
					
			' MsgBox "THIS" & vbCR & "ORIGINAL" & vbCr & hest_original_information & vbCr & vbCr & "CURRENT" & vbCr & hest_current_information
			If hest_current_information <> hest_original_information Then
				hest_updated = true
				Call access_HEST_panel("WRITE", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
			End If
		End If
	End If
end function
Dim shel_updated, hest_updated

function diet_dialog()
	EditBox 395, 0, 45, 15, diet_date_received		
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
	DropListBox 135, 215, 35, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_treatment_plan_dropdown			
	EditBox 270, 215, 55, 15, diet_length_diet
	DropListBox 55, 235, 60, 15, ""+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", diet_status_dropdown 
	CheckBox 125, 235, 195, 10, "Check here to set TIKL for renewal", diet_tikl_checkbox
	EditBox 50, 260, 290, 15, diet_comments
	PushButton 5, 280, 80, 15, "CM23.12- Special Diets", diet_link_CM_special_diet
    PushButton 95, 280, 115, 15, "Processing Special Diet Referrals", diet_SP_referrals
	Text 5, 5, 220, 10, diet_form_name
	Text 340, 5, 55, 10, "Document Date:"
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
end function
Dim diet_date_received, diet_member_number, diet_mfip_msa_status, diet_1_dropdown, diet_2_dropdown, diet_3_dropdown, diet_4_dropdown, diet_5_dropdown, diet_6_dropdown, diet_7_dropdown, diet_8_dropdown, diet_relationship_1_dropdown, diet_relationship_2_dropdown, diet_relationship_3_dropdown, diet_relationship_4_dropdown, diet_relationship_5_dropdown, diet_relationship_6_dropdown, diet_relationship_7_dropdown, diet_relationship_8_dropdown, diet_verif_1_dropdown, diet_verif_2_dropdown, diet_verif_3_dropdown, diet_verif_4_dropdown, diet_verif_5_dropdown, diet_verif_6_dropdown, diet_verif_7_dropdown, diet_verif_8_dropdown, diet_date_last_exam, diet_treatment_plan_dropdown, diet_status_dropdown, diet_length_diet, diet_comments, diet_tikl_checkbox

function other_dialog()
	EditBox 395, 0, 45, 15, other_date_received	
	EditBox 75, 25, 260, 15, other_list_form_names
	EditBox 75, 50, 260, 15, other_doc_notes
	EditBox 75, 75, 260, 15, other_verif_received
	EditBox 75, 100, 260, 15, other_action_taken
	Text 340, 5, 55, 10, "Document Date:"
	Text 25, 30, 50, 10, "Form Name(s)"
	Text 20, 80, 55, 10, "Verifs Received"
	Text 25, 105, 50, 10, "Actions Taken"
	Text 15, 55, 55, 10, "Document Notes"
	Text 5, 5, 220, 10, other_form_name
end function
Dim other_date_received, other_list_form_names, other_doc_notes, other_verif_received, other_action_taken

function dialog_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs. 
	If form_count < Ubound(form_type_array, 2) AND ButtonPressed = -1 Then	ButtonPressed = next_btn	'If the enter button is selected  and we are not at the last dailog, the script will handle this as if Next was selected
	If form_count = Ubound(form_type_array, 2) AND ButtonPressed = -1 Then ButtonPressed = complete_btn	'If the enter button is selected and we are at the last dailog, the script will handle this as if Complete was selected
	If ButtonPressed = next_btn AND err_msg = "" Then form_count = form_count + 1	'If next is selected, it will iterate to the next form in the array and display this dialog
	If ButtonPressed = previous_btn AND err_msg = "" Then form_count = form_count - 1	'If previous is selected, it will iterate to the previous form in the array and display this dialog
	If (ButtonPressed = asset_btn OR ButtonPressed = atr_btn OR ButtonPressed = arep_btn OR ButtonPressed = change_btn OR ButtonPressed = evf_btn OR ButtonPressed = hospice_btn OR ButtonPressed = iaa_btn OR ButtonPressed = ltc_1503_btn OR ButtonPressed = mof_btn OR ButtonPressed = mtaf_btn OR ButtonPressed = psn_btn OR ButtonPressed = sf_btn OR ButtonPressed = diet_btn OR ButtonPressed = other_btn) AND err_msg = "" Then
		For i = 0 to Ubound(form_type_array, 2) 	'For/Next used to iterate through the array to display the correct dialog
			If ButtonPressed = asset_btn and form_type_array(form_type_const, i) = asset_form_name Then form_count = i 
			If ButtonPressed = atr_btn and form_type_array(form_type_const, i) = atr_form_name Then form_count = i 
			If ButtonPressed = arep_btn and form_type_array(form_type_const, i) = arep_form_name Then form_count = i 
			If ButtonPressed = change_btn and form_type_array(form_type_const, i) = change_form_name Then form_count = i 
			If ButtonPressed = evf_btn and form_type_array(form_type_const, i) = evf_form_name Then form_count = i 
			If ButtonPressed = hospice_btn and form_type_array(form_type_const, i) = hosp_form_name Then form_count = i 
			If ButtonPressed = iaa_btn and form_type_array(form_type_const, i) = iaa_form_name Then form_count = i 
			If ButtonPressed = ltc_1503_btn and form_type_array(form_type_const, i) = ltc_1503_form_name Then form_count = i 
			If ButtonPressed = mof_btn and form_type_array(form_type_const, i) = mof_form_name Then form_count = i 
			If ButtonPressed = mtaf_btn and form_type_array(form_type_const, i) = mtaf_form_name Then form_count = i 
			If ButtonPressed = psn_btn and form_type_array(form_type_const, i) = psn_form_name Then form_count = i 
			If ButtonPressed = sf_btn and form_type_array(form_type_const, i) = sf_form_name Then form_count = i 
			If ButtonPressed = diet_btn and form_type_array(form_type_const, i) = diet_form_name Then form_count = i 
			If ButtonPressed = other_btn and form_type_array(form_type_const, i) = other_form_name Then form_count = i 
		Next
	End If 
	'Handling for resrouces 
	If ButtonPressed = hosp_TE0207081_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.07.081%20HOSPICE%20CASES.pdf?csf=1&web=1&e=WgdqsC"
	If ButtonPressed = hosp_SP_hospice_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Hospice.aspx"
	If ButtonPressed = iaa_CM121203_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00121203"
	If ButtonPressed = iaa_te021214_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.12.14%20INTERIM%20ASSISTANCE%20REIMBURSEMENT%20INTERFACE.pdf?csf=1&web=1&e=tUXs96"
	If ButtonPressed = iaa_smi_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://smi.dhs.state.mn.us/login"
	If ButtonPressed = mtaf_cm101801_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00101801"
	If ButtonPressed = mtaf_cm0510_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_000510"
	If ButtonPressed = mtaf_mfip_orientation_info_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MFIP_Orientation.aspx"
	If ButtonPressed = mtaf_cm15121206_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_0005121206"
	If ButtonPressed = diet_link_CM_special_diet Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_002312"
	If ButtonPressed = diet_SP_referrals Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Processing_Special_Diet_Referral.aspx"
	If ButtonPressed = psn_CM1315_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_001315"
	If ButtonPressed = psn_TE1817_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2018.17%20ADULT%20GRH%20BASIS%20OF%20ELIGIBILITY.pdf?csf=1&web=1&e=7YWKmj"
	If ButtonPressed = psn_hss_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=DHS-316637"
	If ButtonPressed = psn_mhm_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg/Training_home_page.doc?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=dhs16_184936#em"
	If ButtonPressed = psn_hsss_btn	Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=dhs-289228"
end function 

function form_specific_error_handling()	'Error handling for main dialog of forms
	If (ButtonPressed = complete_btn OR ButtonPressed = previous_btn OR ButtonPressed = next_btn OR ButtonPressed = -1 OR ButtonPressed = asset_btn OR ButtonPressed = atr_btn OR ButtonPressed = arep_btn OR ButtonPressed = change_btn OR ButtonPressed = evf_btn OR ButtonPressed = hospice_btn OR ButtonPressed = iaa_btn OR ButtonPressed = ltc_1503_btn OR ButtonPressed = mof_btn OR ButtonPressed = mtaf_btn OR ButtonPressed = psn_btn OR ButtonPressed = sf_btn OR ButtonPressed = diet_btn OR ButtonPressed = other_btn) Then 		'Error handling will display at the point of each dialog and will not let the user continue unless the applicable errors are resolved. Had to list all buttons including -1 so ensure the error reporting is called and hit when the script is run.
		For form_errors = 0 to Ubound(form_type_array, 2)
			If form_type_array(form_type_const, form_errors) = asset_form_name AND current_dialog = "asset" then 'Error handling for Asset Form
				actions_taken = Trim(actions_taken)
				actions_taken = actions_taken & ", "
				If IsDate(asset_date_received) = FALSE Then asset_err_msg = asset_err_msg & vbNewLine & "* You must enter a valid date for the Document Date."
				If actions_taken = "" Then asset_err_msg = asset_err_msg & vbNewLine & "* You must enter your actions taken."
				If (asset_dhs_6054_checkbox = checked AND IsDate(asset_date_received) = FALSE) Then asset_err_msg = asset_err_msg & vbNewLine & "* You must enter Document Date."
				Call asset_dialog_DHS6054_and_update_asset_panels		'This will call additional asset dialogs if DHS6054 or update asset panels is checked
			End If

			If form_type_array(form_type_const, form_errors) = atr_form_name AND current_dialog = "atr" Then 'Error handling for ATR Form 
				If IsDate(atr_date_received) = FALSE Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the Document Date."
				If atr_member_dropdown = "Select" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a member from the Member dropdown."
				If IsDate(atr_start_date) = FALSE Then  atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the Start Date."
				If IsDate(atr_end_date) = FALSE Then  atr_err_msg = atr_err_msg & vbNewLine & "* Enter a valid date for the End Date."
				If trim(atr_authorization_type) = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a valid authorization type from the dropdown"
				If trim(atr_contact_type) = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Select a valid contact type from the dropdown"
				If trim(atr_name) = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter contact name"
				If trim(atr_phone_number) = "" Then atr_err_msg = atr_err_msg & vbNewLine & "* Enter phone number"
				If (atr_eval_treat_checkbox = 0 and atr_coor_serv_checkbox = 0 and atr_elig_serv_checkbox = 0 and atr_court_checkbox = 0 and atr_other_checkbox = 0) Then atr_err_msg = atr_err_msg & vbNewLine & "* Must check at least one checkbox indicating use of requested record"
				If (atr_other_checkbox = checked and trim(atr_other) = "") Then err_msg = err_msg & vbNewLine & "* Other checkbox checked, specify details in the box below checkbox"
				If (trim(atr_other) <> "" and atr_other_checkbox = unchecked) Then atr_err_msg = atr_err_msg & vbNewLine & "* Other text field must be blank unless Other checkbox is checked"
			End If

			If form_type_array(form_type_const, form_errors) = arep_form_name AND current_dialog = "arep" then 'Error handling for AREP Form 
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
				If IsDate(AREP_recvd_date) = False Then arep_err_msg = arep_err_msg & vbNewLine & "* Enter the date the form was received."
				IF (arep_SNAP_AREP_checkbox <> checked AND arep_HC_AREP_checkbox <> checked AND arep_CASH_AREP_checkbox <> checked) THEN arep_err_msg = arep_err_msg & vbNewLine &"* Select a program"
				IF isdate(arep_signature_date) = false THEN arep_err_msg = arep_err_msg & vbNewLine & "* Enter a valid date for the date the form was signed/valid from."
				IF (arepTIKL_check = checked AND trim(arep_signature_date) = "") THEN arep_err_msg = arep_err_msg & vbNewLine & "* You have requested the script to TIKL based on the signature date but you did not enter the signature date."
			End If

			If form_type_array(form_type_const, form_errors) = change_form_name AND current_dialog = "chng" then 'Error handling for Change Form 
				If IsDate(chng_effective_date) = False Then chng_err_msg = chng_err_msg & vbNewLine & "* Enter a valid Effective date."
				If IsDate(chng_date_received) = False Then chng_err_msg = chng_err_msg & vbNewLine & "* Enter a valid date Document received date."  ' Validate that Date Change Reported/Received field is not empty and is in a proper date format
				If trim(chng_address_notes) = "" AND trim(chng_household_notes) = "" AND trim(chng_asset_notes) = "" AND trim(chng_vehicles_notes) = "" AND trim(chng_income_notes) = "" AND trim(chng_shelter_notes) = "" AND trim(chng_other_change_notes) = "" THEN chng_err_msg = chng_err_msg & vbNewLine & "* All change reported fields are blank. At least one needs info."  ' Validate the Changes Reported fields to ensure that at least one field is filled in
				If trim(chng_actions_taken) = "" AND trim(chng_other_notes) = "" AND trim(chng_verifs_requested) = "" THEN chng_err_msg = chng_err_msg & vbNewLine & "* All of the Actions fields are blank. At least one need info."  ' Validate the Actions fields to ensure that at least one field is filled in
				If trim(chng_notable_change) = "" Then chng_err_msg = chng_err_msg & vbNewLine & "* Notable changes reported is blank, make a selection."
				If chng_changes_continue = "Select One:" THEN chng_err_msg = chng_err_msg & vbNewLine & "* Indicate whether changes will or will not continue next month."  ' Validate that worker selects option from dropdown list as to how long change will last
			End If

			If form_type_array(form_type_const, form_errors) = evf_form_name AND current_dialog = "evf" then 'Error handling for EVF Form 
				IF IsDate(evf_date_received) = FALSE THEN evf_err_msg = evf_err_msg & vbCr & "* Enter a vaid Document Date."
				If EVF_status_dropdown = "Select one..." THEN evf_err_msg = evf_err_msg & vbCr & "* Select the status of the EVF on the dropdown menu"		'checks that there is a date in the date received box
				IF trim(evf_employer) = "" THEN evf_err_msg = evf_err_msg & vbCr & "* Enter the employers name."  'checks if the employer name has been entered
				IF evf_client = "Select" THEN evf_err_msg = evf_err_msg & vbCr & "* Enter the MEMB information."  'checks if the client name has been entered
				IF evf_info = "Select one..." THEN evf_err_msg = evf_err_msg & vbCr & "* Select if additional info was requested."  'checks if completed by employer was selected
				IF evf_info = "yes" and IsDate(evf_info_date) = FALSE THEN evf_err_msg = evf_err_msg & vbCr & "* Enter a valid date that additional info was requested."  'checks that there is a info request date entered if the it was requested
				IF evf_info = "yes" and evf_request_info = "" THEN evf_err_msg = evf_err_msg & vbCr & "* Enter the method used to request additional info."		'checks that there is a method of inquiry entered if additional info was requested
				If evf_info = "no" and evf_request_info <> "" then evf_err_msg = evf_err_msg & vbCr & "* You cannot mark additional info as 'no' and have information requested."
				If evf_info = "no" and evf_info_date <> "" then evf_err_msg = evf_err_msg & vbCr & "* You cannot mark additional info as 'no' and have a date requested."
				If EVF_TIKL_checkbox = 1 and evf_info <> "yes" then evf_err_msg = evf_err_msg & vbCr & "* Additional information was not requested, uncheck the TIKL checkbox."
			End If

			If form_type_array(form_type_const, form_errors) = hosp_form_name AND current_dialog = "hosp" then 'Error handling for Hospice Form 
				If IsDate(hosp_date_received) = FALSE Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter a valid date for the Document Date." 
				If hosp_resident_name = "Select" Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Select the resident that is in hospice."
				If trim(hosp_name) = "" Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter the name of the Hospice the client entered."       'hospice name required
				If IsDate(hosp_entry_date) = FALSE Then hosp_err_msg = hosp_err_msg & vbNewLine & "* Enter a valid date for the Hospice Entry."   'entry date also required
			End If

			If form_type_array(form_type_const, form_errors) = iaa_form_name AND current_dialog = "iaa" then 'Error handling for IAA Form 
				IF IsDate(iaa_date_received) = FALSE THEN iaa_err_msg = iaa_err_msg & vbCr & "* Enter a valid Document date."
				If iaa_member_dropdown = "Select" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Select the resident from the dropdown."
				If iaa_form_received_checkbox = unchecked and iaa_ssi_form_received_checkbox = unchecked Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Must select which type(s) of IAA received"
				If iaa_form_received_checkbox = Checked and iaa_type_assistance = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Select Type of interim assistance for IAA"
				If iaa_ssi_form_received_checkbox = Checked and iaa_ssi_type_assistance = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Select AG or HS interim assistance for IAA-SSI"
				If iaa_update_pben_checkbox = checked Then 	'Only if they select saying they want to update PBEN will the following fields be required
					If (iaa_benefit_type = "" or trim(iaa_referral_date) = "" or iaa_disposition_code_dropdown = "") Then 	'Only requiring fields that are required in pben panel to save. 
						iaa_err_msg = iaa_err_msg & vbNewLine & "* PBEN field requirements:"
						If iaa_benefit_type = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Select benefit type"
						If iaa_disposition_code_dropdown = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Select Disposition Code"
						If IsDate(iaa_referral_date) = FALSE Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter a valid Referral Date"
						If iaa_date_applied_pben <> "" AND IsDate(iaa_date_applied_pben) = FALSE Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter a valid Applied to PBEN date"
						If iaa_iaa_date <> "" AND IsDate(iaa_iaa_date) = FALSE Then iaa_err_msg = iaa_err_msg & vbNewLine & "  * Enter a valid IAA date"
					End If
				End If
				If iaa_update_pben_checkbox = unchecked AND iaa_comments = "" Then iaa_err_msg = iaa_err_msg & vbNewLine & "* Must explain in comments why PBEN is not being created/updated. "
			End If

			If form_type_array(form_type_const, form_errors) = ltc_1503_form_name AND current_dialog = "ltc 1503" then 'Error handling for LTC 1503 Form
				If IsDate(ltc_1503_date_received) = FALSE THEN ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "* Enter a valid Document date."
				If IsDate(ltc_1503_admit_date) = FALSE Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "* Enter valid admission date"
				If ltc_1503_discharge_date <> "" AND IsDate(ltc_1503_discharge_date) = FALSE Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "* Enter valid discharge date"
				If ltc_1503_FACI_update_checkbox = checked AND (trim(ltc_1503_FACI_1503) = "" OR ltc_1503_length_of_stay = "" OR ltc_1503_level_of_care = "" OR trim(ltc_1503_admit_date) = "" OR trim(ltc_1503_faci_footer_month) = "" OR trim(ltc_1503_faci_footer_year) = "") Then 
					ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "* Update FACI Panel selected. Complete the required fields:"
					If trim(ltc_1503_FACI_1503) = "" Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "-   Enter facility name"
					If ltc_1503_length_of_stay = "" Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "-   Select length of stay"
					If ltc_1503_level_of_care = "" Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "-   Select level of care"
					If IsDate(ltc_1503_admit_date) = FALSE Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "-   Enter vaid admission date"
					IF IsNumeric(ltc_1503_faci_footer_month) = FALSE OR IsNumeric(ltc_1503_faci_footer_year) = FALSE THEN ltc_1503_err_msg = ltc_1503_err_msg & vbNewLine &  "-   Enter valid FACI footer month and year."
				End If
				If ltc_1503_sent_verif_request_checkbox = checked AND trim(ltc_1503_sent_request_to) = "" Then ltc_1503_err_msg = ltc_1503_err_msg & vbCr & "* Select/Enter verif sent to"
			End If

			If form_type_array(form_type_const, form_errors) = mof_form_name AND current_dialog = "mof" then 'Error handling for MOF Form 
				If IsDate(mof_date_received) = FALSE Then mof_err_msg = mof_err_msg & vbNewLine & "* Enter a valid Document date."
				If mof_hh_memb = "Select" Then mof_err_msg = mof_err_msg & vbNewLine & "* Select the member from the dropdown."
				IF trim(mof_actions_taken) = "" THEN mof_err_msg = mof_err_msg & vbCr & "* Enter your actions taken."		'checks that notes were entered
				If MOF_TTL_email_checkbox = checked Then
					If IsDate(mof_TTL_email_date) = FALSE Then mof_err_msg = mof_err_msg & vbNewLine & "* Enter a valid date for the date an email about this MOF was sent to TTL."
				End If
				mof_last_exam_date = trim(mof_last_exam_date)
				mof_doctor_date = trim(mof_doctor_date)
				If mof_time_condition_will_last = "Select or Type" Then mof_time_condition_will_last = ""
				mof_time_condition_will_last = trim(mof_time_condition_will_last)
				mof_ability_to_work = trim(mof_ability_to_work)
				mof_other_notes = trim(mof_other_notes)
			End If

			If form_type_array(form_type_const, form_errors) = mtaf_form_name AND current_dialog = "mtaf" then 'Error handling for MTAF Form
				If IsDate(MTAF_date) = False Then mtaf_err_msg = mtaf_err_msg & vbNewLine & "* Enter the date the MTAF was received."
				If MTAF_status_dropdown = "Select one..." Then mtaf_err_msg = mtaf_err_msg & vbNewLine & "* Indicate the status of the MTAF."
				If mtaf_sub_housing_droplist = "Select one..." Then mtaf_err_msg = mtaf_err_msg & vbNewLine & "* Indicate if housing is subsidized or not."
			End If

			If form_type_array(form_type_const, form_errors) = psn_form_name AND current_dialog = "psn" then 'Error handling for PSN Form
				IF IsDate(psn_date_received) = FALSE THEN psn_err_msg = psn_err_msg & vbCr & "* Enter a valid Document Date."
				If psn_member_dropdown = "Select" Then psn_err_msg = psn_err_msg & vbNewLine & "* Select the resident from the dropdown."
				If psn_section_1_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 1 make selection from dropdown."
				If psn_section_2_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 2 make selection from dropdown."
				If psn_section_3_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 3 make selection from dropdown."
				If psn_section_4_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 4 make selection from dropdown."
				If psn_section_5_dropdown = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* For Section 5 make selection from dropdown."
				If trim(psn_cert_prof) = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Certified Professional or NA"
				If trim(psn_facility) = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Enter Facilty name or NA"
				If psn_udpate_wreg_disa_checkbox = checked Then 
					If psn_wreg_fs_pwe = "" OR psn_wreg_work_wreg_status = "" OR psn_wreg_abawd_status = "" OR psn_wreg_ga_elig_status = "" OR trim(psn_disa_begin_date) = "" OR psn_disa_status = "" OR psn_disa_verif = "" Then psn_err_msg = psn_err_msg & vbNewLine & "* Update WREG/DISA checked - complete the required fields:"
					If psn_wreg_fs_pwe = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select FS PWE from dropdown"
					If psn_wreg_work_wreg_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select FSET Work Reg Status from dropdown"
					If psn_wreg_abawd_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select ABAWD Status from dropdown"
					If psn_wreg_ga_elig_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select GA Elig Basis Code from dropdown"
					If IsDate(psn_disa_begin_date) = FALSE Then psn_err_msg = psn_err_msg & vbNewLine & "  * Enter Disa Begin date"
					If psn_disa_status = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select Disa Status from dropdown"
					If psn_disa_verif = "" Then psn_err_msg = psn_err_msg & vbNewLine & "  * Select Verification from dropdown"
				End If
			End If

			If form_type_array(form_type_const, form_errors) = sf_form_name AND current_dialog = "sf" then 'Error handling for Shelter Form
				IF IsDate(sf_date_received) = FALSE THEN sf_err_msg = sf_err_msg & vbCr & "* Enter a valid Document Date." 
				If sf_other_checkbox = checked AND trim(sf_other_utility) = "" Then sf_err_msg = sf_err_msg & vbNewLine & "* You checked 'other', therefore you must specify specific utility."
				If sf_other_checkbox = unchecked AND trim(sf_other_utility) <> "" Then sf_err_msg = sf_err_msg & vbNewLine & "* You specified 'other' utility,'other checkbox' must be checked if field is filled out."
				Call addr_shel_hest_panel_dialog
			End If

			If form_type_array(form_type_const, form_errors) = diet_form_name AND current_dialog = "diet" then 'Error handling for Diet Form 
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
				
				'TODO: Havent' tested on MFIP case handling to limit to 2 diets. Consider hiding extra boxes on dialog
				If (diet_mfip_msa_status = "MFIP-Active") OR (diet_mfip_msa_status = "MFIP-Pending") Then
					'MsgBox "diet_mfip_msa_status" & diet_mfip_msa_status
						If diet_3_dropdown <>"" OR diet_4_dropdown <>"" OR diet_5_dropdown <>"" OR diet_6_dropdown <>"" OR diet_7_dropdown <>"" OR diet_8_dropdown <>"" Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have more than 2 diets for MFIP cases"
				End If

				If Instr(all_diet_string, "*01-High Protein*") AND Instr(all_diet_string, "*02-Controlled protein 40-60 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."
				If Instr(all_diet_string, "*01-High Protein*") AND Instr(all_diet_string,"*03-Controlled protein <40 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."
				If Instr(all_diet_string, "*02-Controlled protein 40-60 grams*") AND Instr(all_diet_string,"*03-Controlled protein <40 grams*") Then diet_err_msg = diet_err_msg & vbNewLine & "* Cannot have multiple protien diets."

				If IsDate(diet_date_last_exam) = FALSE Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter a valid date for Date of last exam."
				If diet_treatment_plan_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select dropdown indicating person is following treatment plan"
				If trim(diet_length_diet) = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Enter length of prescribed diet"
				If diet_status_dropdown = "" Then diet_err_msg = diet_err_msg & vbNewLine & "* Select Diet Status"
				If diet_status_dropdown = "Denied" AND diet_comments = "" Then diet_err_msg = diet_err_msg & vbNewLine & "*Diet Denied, state reason for ineligibility & benefit end date in Comments"
			End If

			If form_type_array(form_type_const, form_errors) = other_form_name AND current_dialog = "other" then 'Error handling for Other Form 
				IF IsDate(other_date_received) = FALSE THEN other_err_msg = other_err_msg & vbCr & "* Enter a valid Document Date."
				If Trim(other_list_form_names) = ""  THEN other_err_msg = other_err_msg & vbCr & "* Specify name of form(s)"
				If Trim(other_doc_notes) = "" THEN other_err_msg = other_err_msg & vbCr &  "* Specify details for each form"
				If Trim(other_action_taken) = "" THEN other_err_msg = other_err_msg & vbCr &  "* Specify actions taken"
			End If
		Next
	End If

	'Complete button triggers the error message to populate. Formatting error meessage to: Adds headers for each form if there are applicable errors
	If asset_err_msg <> "" AND current_dialog = "asset" Then err_msg = err_msg & vbNewLine & "ASSET DIALOG" & asset_err_msg & vbNewLine
	If atr_err_msg <> "" AND current_dialog = "atr" Then err_msg = err_msg & vbNewLine & "ATR DIALOG" & atr_err_msg & vbNewLine
	If arep_err_msg <> "" AND current_dialog = "arep" Then err_msg = err_msg & vbNewLine & "AREP DIALOG" & arep_err_msg & vbNewLine
	If chng_err_msg <> "" AND current_dialog = "chng" Then err_msg = err_msg & vbNewLine & "CHANGE DIALOG" & chng_err_msg & vbNewLine
	If evf_err_msg <> "" AND current_dialog = "evf" Then err_msg = err_msg & vbNewLine & "EVF DIALOG" & evf_err_msg & vbNewLine
	If hosp_err_msg <> "" AND current_dialog = "hosp" Then err_msg = err_msg & vbNewLine & "HOSPICE DIALOG" & hosp_err_msg & vbNewLine
	If iaa_err_msg <> "" AND current_dialog = "iaa" Then err_msg = err_msg & vbNewLine & "IAA DIALOG" & iaa_err_msg & vbNewLine
	If ltc_1503_err_msg <> "" AND current_dialog = "ltc 1503" Then err_msg = err_msg & vbNewLine & "LTC 1503 DIALOG" & ltc_1503_err_msg & vbNewLine
	If mof_err_msg <> "" AND current_dialog = "mof" Then err_msg = err_msg & vbNewLine & "MOF DIALOG" & mof_err_msg & vbNewLine
	If mtaf_err_msg <> "" AND current_dialog = "mtaf" Then err_msg = err_msg & vbNewLine & "MTAF DIALOG" & mtaf_err_msg & vbNewLine
	If psn_err_msg <> "" AND current_dialog = "psn" Then err_msg = err_msg & vbNewLine & "PSN DIALOG" & psn_err_msg & vbNewLine
	If sf_err_msg <> "" AND current_dialog = "sf" Then err_msg = err_msg & vbNewLine & "SF DIALOG" & sf_err_msg & vbNewLine
	If diet_err_msg <> "" AND current_dialog = "diet" Then err_msg = err_msg & vbNewLine & "DIET DIALOG" & diet_err_msg & vbNewLine
	If other_err_msg <> "" AND current_dialog = "other" Then err_msg = err_msg & vbNewLine & "OTHER FORM DIALOG" & other_err_msg & vbNewLine

	'If complete button or enter while on last tab is selected and all forms are not complete, this will stop them from proceeding by listing the outstanding forms as an error message
	If ButtonPressed = complete_btn OR (form_count = Ubound(form_type_array, 2) AND ButtonPressed = -1) Then 
		For item = 0 to Ubound(form_type_array, 2)
			If item = form_type_array(the_last_const, item) AND asset_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Asset form not complete~~"
			If item = form_type_array(the_last_const, item) AND atr_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~ATR form not complete~~"
			If item = form_type_array(the_last_const, item) AND arep_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~AREP form not complete~~"
			If item = form_type_array(the_last_const, item) AND chng_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Change form not complete~~"
			If item = form_type_array(the_last_const, item) AND evf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~EVF form not complete~~"
			If item = form_type_array(the_last_const, item) AND hosp_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Hospice form not complete~~"
			If item = form_type_array(the_last_const, item) AND iaa_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~IAA form not complete~~"
			If item = form_type_array(the_last_const, item) AND ltc_1503_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~LTC 1503 form not complete~~"
			If item = form_type_array(the_last_const, item) AND mof_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~MOF form not complete~~"
			If item = form_type_array(the_last_const, item) AND mtaf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~MTAF form not complete~~"
			If item = form_type_array(the_last_const, item) AND psn_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~PSN form not complete~~"
			If item = form_type_array(the_last_const, item) AND sf_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Proof of Shelter/Residence form not complete~~"
			If item = form_type_array(the_last_const, item) AND diet_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Diet form not complete~~"
			If item = form_type_array(the_last_const, item) AND other_err_msg <> "" Then err_msg = err_msg & vbNewLine & "~~Other form not complete~~"
		Next 
	End If

	If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
end function

'Check for case number & footer & background
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'DIALOG COLLECTING CASE, FOOTER MO/YR===========================================================================
Do
	DO
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 196, 105, "NOTES - Docs Received " & MAXIS_case_number
			EditBox 70, 5, 50, 15, MAXIS_case_number
			EditBox 70, 25, 20, 15, MAXIS_footer_month
			EditBox 100, 25, 20, 15, MAXIS_footer_year
			EditBox 70, 45, 100, 15, worker_signature
			ButtonGroup ButtonPressed
				PushButton 140, 5, 50, 15, "Instructions", msg_show_instructions_btn	
				OkButton 85, 85, 50, 15
				CancelButton 140, 85, 50, 15
			Text 20, 10, 50, 10, "Case number: "
			Text 20, 30, 45, 10, "Footer month:"
			Text 5, 50, 60, 10, "Worker signature:"
			Text 5, 70, 185, 10, "Script Purpose: Case note details of documents received"
		EndDialog


		dialog Dialog1	'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		'IF IsNumeric(MAXIS_footer_month) = FALSE OR IsNumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbNewLine &  "* You must type a valid footer month and year."
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		If ButtonPressed = msg_show_instructions_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20DOCUMENTS%20RECEIVED.docx?d=w1dce0cc33ca541f68855f406a63ab02b&csf=1&web=1&e=LXojaV"	'copy the instructions URL here
			err_msg = "LOOP"
		End If
		IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Checking for PRIV cases.
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = TRUE then script_end_procedure("PRIV case, cannot access/update. The script will now end.")

Call back_to_SELF
continue_in_inquiry = ""
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If

Call MAXIS_background_check
Call Generate_Client_List(HH_Memb_DropDown, "Select")         'filling the dropdown with ALL of the household members
CALL Generate_Client_List(client_dropdown, "Select One...")	
CALL Generate_Client_List(client_dropdown_CB, "Select or Type")

'DIALOGS COLLECTING FORM SELECTION===========================================================================
Do							'Do Loop to cycle through dialog as many times as needed until all desired forms are added
	Do
		Do
			err_msg = ""
			Dialog1 = "" 			'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 296, 235, "Select Documents Received"
				DropListBox 30, 30, 180, 15, ""+chr(9)+asset_form_name+chr(9)+atr_form_name+chr(9)+arep_form_name+chr(9)+change_form_name+chr(9)+evf_form_name+chr(9)+hosp_form_name+chr(9)+iaa_form_name+chr(9)+ltc_1503_form_name+chr(9)+mof_form_name+chr(9)+mtaf_form_name+chr(9)+psn_form_name+chr(9)+sf_form_name+chr(9)+diet_form_name+chr(9)+other_form_name, Form_type
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
				End If 
			End If
			'MsgBox "all form array string" & all_form_array '= split(all_form_array, "*")
				
			If ButtonPressed = clear_button Then 'Clear button wipes out any selections already made so the user can reselect correct forms.
				ReDim form_type_array(the_last_const, form_count)		
				form_count = 0							'Reset the form count to 0 so that y_pos resets to 95. 
				Form_string = ""						'Reset string to nothing
				add_to_array = ""						'reset to nothing
				all_form_array = "*"					'Reset string to *		
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
                form_count = 0							'Resetting the form count to 0 so that y_pos resets to 95. 
				Form_string = ""						'Resetting string to nothing 
				all_form_array = "*"					'Resetting list of strings to * 
				add_to_array = ""
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
					CheckBox 15, 90, 160, 10, ltc_1503_form_name, ltc_1503_checkbox
					CheckBox 15, 100, 160, 10, mof_form_name, mof_checkbox
					CheckBox 15, 110, 160, 10, mtaf_form_name, mtaf_checkbox
					CheckBox 15, 120, 160, 10, psn_form_name, psn_checkbox
					CheckBox 15, 130, 160, 10, sf_form_name, shelter_checkbox
					CheckBox 15, 140, 170, 10, diet_form_name, diet_checkbox
					CheckBox 15, 150, 170, 10, other_form_name, other_checkbox
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
				If other_checkbox = checked Then 
					ReDim Preserve form_type_array(the_last_const, form_count)		'ReDim Preserve to keep all selections without writing over one another.
					form_type_array(form_type_const, form_count) = other_form_name
					form_count= form_count + 1 
				End If
			
				'MsgBox "all form array string" & all_form_array 
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and ltc_1503_checkbox = unchecked and mof_checkbox = unchecked and mtaf_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked and other_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""	
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If		
Loop Until ButtonPressed = Ok

'MAXIS NAVIGATION READ===========================================================================
For maxis_panel_read = 0 to Ubound(form_type_array, 2)
	'ASSET CODE-START
	If form_type_array(form_type_const, maxis_panel_read) = asset_form_name Then 'MAXIS NAVIGATION FOR ASSET: READ ACCT, SECU, CARS
		Call HH_member_custom_dialog(HH_member_array)	'This will be for any functionality that needs the HH Member array
		asset_counter = 0
		skip_asset = FALSE
		Do 
			Call navigate_to_MAXIS_screen("STAT", "ACCT")
			EMReadScreen nav_check, 4, 2, 44
			EmWaitReady 0, 0
		Loop until nav_check = "ACCT"

		For each member in HH_member_array
			Call write_value_and_transmit(member, 20, 76)
			EMReadScreen acct_versions, 1, 2, 78
			If acct_versions <> "0" Then
				EMWriteScreen "01", 20, 79
				transmit
				Do
					EMReadScreen ACCT_instance, 1, 2, 73
					EMReadScreen ACCT_type, 2, 6, 44
					EMReadScreen ACCT_nbr, 20, 7, 44
					EMReadScreen ACCT_location, 20, 8, 44
					EMReadScreen ACCT_balance, 8, 10, 46
					EMReadScreen ACCT_bal_verif, 1, 10, 64
					EMReadScreen ACCT_bal_date, 8, 11, 44
					EMReadScreen ACCT_withdraw_pen, 8, 12, 46
					EMReadScreen ACCT_withdraw_YN, 1, 12, 64
					EMReadScreen ACCT_withdraw_verif, 1, 12, 72
					EMReadScreen ACCT_cash, 1, 14, 50
					EMReadScreen ACCT_snap, 1, 14, 57
					EMReadScreen ACCT_hc, 1, 14, 64
					EMReadScreen ACCT_grh, 1, 14, 72
					EMReadScreen ACCT_ive, 1, 14, 80
					EMReadScreen ACCT_joint_owner_YN, 1, 15, 44
					EMReadScreen ACCT_share_ratio, 5, 15, 76
					EMReadScreen ACCT_next_interest, 5, 17, 57
					EMReadScreen ACCT_updated_date, 8, 21, 55

					ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

					ASSETS_ARRAY(ast_panel, asset_counter) = "ACCT"
					ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
					For each person in client_list_array
						If left(person, 2) = member then
							ASSETS_ARRAY(ast_owner, asset_counter) = person
							Exit For
						End If
					Next
					ASSETS_ARRAY(ast_instance, asset_counter) = "0" & ACCT_instance
					If ACCT_type = "SV" Then ASSETS_ARRAY(ast_type, asset_counter) = "SV - Savings"
					If ACCT_type = "CK" Then ASSETS_ARRAY(ast_type, asset_counter) = "CK - Checking"
					If ACCT_type = "CD" Then ASSETS_ARRAY(ast_type, asset_counter) = "CD - Cert of Deposit"
					If ACCT_type = "MM" Then ASSETS_ARRAY(ast_type, asset_counter) = "MM - Money market"
					If ACCT_type = "DC" Then ASSETS_ARRAY(ast_type, asset_counter) = "DC - Debit Card"
					If ACCT_type = "KO" Then ASSETS_ARRAY(ast_type, asset_counter) = "KO - Keogh Account"
					If ACCT_type = "FT" Then ASSETS_ARRAY(ast_type, asset_counter) = "FT - Federatl Thrift SV plan"
					If ACCT_type = "SL" Then ASSETS_ARRAY(ast_type, asset_counter) = "SL - Stat/Local Govt Ret"
					If ACCT_type = "RA" Then ASSETS_ARRAY(ast_type, asset_counter) = "RA - Employee Ret Annuities"
					If ACCT_type = "NP" Then ASSETS_ARRAY(ast_type, asset_counter) = "NP - Non-Profit Employer Ret Plan"
					If ACCT_type = "IR" Then ASSETS_ARRAY(ast_type, asset_counter) = "IR - Indiv Ret Acct"
					If ACCT_type = "RH" Then ASSETS_ARRAY(ast_type, asset_counter) = "RH - Roth IRA"
					If ACCT_type = "FR" Then ASSETS_ARRAY(ast_type, asset_counter) = "FR - Ret Plans for Employers"
					If ACCT_type = "CT" Then ASSETS_ARRAY(ast_type, asset_counter) = "CT - Corp Ret Trust"
					If ACCT_type = "RT" Then ASSETS_ARRAY(ast_type, asset_counter) = "RT - Other Ret Fund"
					If ACCT_type = "QT" Then ASSETS_ARRAY(ast_type, asset_counter) = "QT - Qualified Tuition (529)"
					If ACCT_type = "CA" Then ASSETS_ARRAY(ast_type, asset_counter) = "CA - Coverdell SV (530)"
					If ACCT_type = "OE" Then ASSETS_ARRAY(ast_type, asset_counter) = "OE - Other Educational "
					If ACCT_type = "OT" Then ASSETS_ARRAY(ast_type, asset_counter) = "OT - Other"
					ASSETS_ARRAY(ast_number, asset_counter) = replace(ACCT_nbr, "_", "")
					ASSETS_ARRAY(ast_location, asset_counter) = replace(ACCT_location, "_", "")
					ASSETS_ARRAY(ast_balance, asset_counter) = trim(ACCT_balance)
					If ACCT_bal_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1 - Bank Statement"
					If ACCT_bal_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - Agcy Ver Form"
					If ACCT_bal_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - Coltrl Document"
					If ACCT_bal_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
					If ACCT_bal_verif = "6" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"
					If ACCT_bal_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prvd"
					ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(ACCT_bal_date, " ", "/")
					If ASSETS_ARRAY(ast_bal_date, asset_counter) = "__/__/__" Then ASSETS_ARRAY(ast_bal_date, asset_counter) = ""
					ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(replace(ACCT_withdraw_pen, "_", ""))
					ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = replace(ACCT_withdraw_YN, "_", "")
					ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = replace(ACCT_withdraw_verif, "_", "")
					ASSETS_ARRAY(apply_to_CASH, asset_counter) = replace(ACCT_cash, "_", "")
					ASSETS_ARRAY(apply_to_SNAP, asset_counter) = replace(ACCT_snap, "_", "")
					ASSETS_ARRAY(apply_to_HC, asset_counter) = replace(ACCT_hc, "_", "")
					ASSETS_ARRAY(apply_to_GRH, asset_counter) = replace(ACCT_grh, "_", "")
					ASSETS_ARRAY(apply_to_IVE, asset_counter) = replace(ACCT_ive, "_", "")
					ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = replace(ACCT_joint_owner_YN, "_", "")
					ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(ACCT_share_ratio, " ", "")
					ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = replace(ACCT_next_interest, " ", "/")
					If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = "__/__" Then ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = ""
					ASSETS_ARRAY(update_panel, asset_counter) = unchecked
					ASSETS_ARRAY(update_date, asset_counter) = replace(ACCT_updated_date, " ", "/")
					transmit
					asset_counter = asset_counter + 1
					EMReadScreen reached_last_ACCT_panel, 13, 24, 2
				Loop until reached_last_ACCT_panel = "ENTER A VALID"
			End If
		Next
		
		Do
			Call navigate_to_MAXIS_screen("STAT", "SECU")
			EMReadScreen nav_check, 4, 2, 45
			EmWaitReady 0, 0
		Loop until nav_check = "SECU"

		For each member in HH_member_array
			Call write_value_and_transmit(member, 20, 76)

			EMReadScreen secu_versions, 1, 2, 78
			If secu_versions <> "0" Then
				EMWriteScreen "01", 20, 79
				transmit
				Do

					EMReadScreen SECU_instance, 1, 2, 73
					EMReadScreen SECU_type, 2, 6, 50
					EMReadScreen SECU_acct_number, 12, 7, 50
					EMReadScreen SECU_name, 20, 8, 50
					EMReadScreen SECU_csv, 8, 10, 52
					EMReadScreen SECU_value_date, 8, 11, 35
					EMReadScreen SECU_verif, 1, 11, 50
					EMReadScreen SECU_face_value, 8, 12, 52
					EMReadScreen SECU_withdraw_amount, 8, 13, 52
					EMReadScreen SECU_wthdrw_YN, 1, 13, 72
					EMReadScreen SECU_wthdrw_verif, 1, 13, 80
					EMReadScreen SECU_apply_to_CASH, 1, 15, 50
					EMReadScreen SECU_apply_to_SNAP, 1, 15, 57
					EMReadScreen SECU_apply_to__HC, 1, 15, 64
					EMReadScreen SECU_apply_to_GRH, 1, 15, 72
					EMReadScreen SECU_apply_to_IVE, 1, 15, 80
					EMReadScreen SECU_joint_owner_YN, 1, 16, 44
					EMReadScreen SECU_share_ratio, 5, 16, 76
					EMReadScreen SECU_updated_date, 8, 21, 55


					ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

					ASSETS_ARRAY(ast_panel, asset_counter) = "SECU"
					ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
					For each person in client_list_array
						If left(person, 2) = member then
							ASSETS_ARRAY(ast_owner, asset_counter) = person
							Exit For
						End If
					Next
					ASSETS_ARRAY(ast_instance, asset_counter) = "0" & SECU_instance
					If SECU_type = "LI" Then ASSETS_ARRAY(ast_type, asset_counter) = "LI - Life Insurance"
					If SECU_type = "ST" Then ASSETS_ARRAY(ast_type, asset_counter) = "ST - Stocks"
					If SECU_type = "BO" Then ASSETS_ARRAY(ast_type, asset_counter) = "BO - Bonds"
					If SECU_type = "CD" Then ASSETS_ARRAY(ast_type, asset_counter) = "CD - Ctrct For Deed"
					If SECU_type = "MO" Then ASSETS_ARRAY(ast_type, asset_counter) = "MO - Mortgage Note"
					If SECU_type = "AN" Then ASSETS_ARRAY(ast_type, asset_counter) = "AN - Annuity"
					If SECU_type = "OT" Then ASSETS_ARRAY(ast_type, asset_counter) = "OT - Other"
					ASSETS_ARRAY(ast_number, asset_counter) = replace(SECU_acct_number, "_", "")
					ASSETS_ARRAY(ast_location, asset_counter) = replace(SECU_name, "_", "")
					ASSETS_ARRAY(ast_csv, asset_counter) = trim(SECU_csv)
					ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(SECU_value_date, " ", "/")
					If SECU_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1  - Agency Form"
					If SECU_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - Source Doc"
					If SECU_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - Phone Contact"
					If SECU_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
					If SECU_verif = "6" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Stmt"
					If SECU_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prov"
					ASSETS_ARRAY(ast_face_value, asset_counter) = replace(trim(SECU_face_value), "_", "")
					ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(replace(SECU_withdraw_amount, "_", ""))
					ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = replace(SECU_wthdrw_YN, "_", "")
					ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = replace(SECU_wthdrw_verif, "_", "")
					ASSETS_ARRAY(apply_to_CASH, asset_counter) = replace(SECU_apply_to_CASH, "_", "")
					ASSETS_ARRAY(apply_to_SNAP, asset_counter) = replace(SECU_apply_to_SNAP, "_", "")
					ASSETS_ARRAY(apply_to_HC, asset_counter) = replace(SECU_apply_to_HC, "_", "")
					ASSETS_ARRAY(apply_to_GRH, asset_counter) = replace(SECU_apply_to_GRH, "_", "")
					ASSETS_ARRAY(apply_to_IVE, asset_counter) = replace(SECU_apply_to_IVE, "_", "")
					ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = replace(SECU_joint_owner_YN, "_", "")
					ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(SECU_share_ratio, " ", "")
					ASSETS_ARRAY(update_date, asset_counter) = replace(SECU_updated_date, " ", "/")
					ASSETS_ARRAY(update_panel, asset_counter) = Unchecked

					transmit
					asset_counter = asset_counter + 1
					EMReadScreen reached_last_SECU_panel, 13, 24, 2
				Loop until reached_last_SECU_panel = "ENTER A VALID"
			End If
		Next
		
		Do
			Call navigate_to_MAXIS_screen("STAT", "CARS")
			EMReadScreen nav_check, 4, 2, 44
			EmWaitReady 0, 0
		Loop until nav_check = "CARS"
		For each member in HH_member_array
			Call write_value_and_transmit(member, 20, 76)

			EMReadScreen cars_versions, 1, 2, 78
			If cars_versions <> "0" Then
				EMWriteScreen "01", 20, 79
				transmit
				Do

					EMReadScreen CARS_instance, 1, 2, 73
					EMReadScreen CARS_type, 1, 6, 43
					EMReadScreen CARS_year, 4, 8, 31
					EMReadScreen CARS_make, 15, 8, 43
					EMReadScreen CARS_model, 15, 8, 66
					EMReadScreen CARS_trade_in, 8, 9, 45
					EMReadScreen CARS_loan, 8, 9, 62
					EMReadScreen CARS_source, 1, 9, 80
					EMReadScreen CARS_owner_verif, 1, 10, 60
					EMReadScreen CARS_owe_amount, 8, 12, 45
					EMReadScreen CARS_owed_verif, 1, 12, 60
					EMReadScreen CARS_owed_date, 8, 13, 43
					EMReadScreen CARS_use, 1, 15, 43
					EMReadScreen CARS_hc_benefit, 1, 15, 76
					EMReadScreen CARS_joint_owner_YN, 1, 16, 43
					EMReadScreen CARS_share_ratio, 5, 16, 76
					EMReadScreen CARS_updated_date, 8, 21, 55

					ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

					ASSETS_ARRAY(ast_panel, asset_counter) = "CARS"
					ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
					For each person in client_list_array
						If left(person, 2) = member then
							ASSETS_ARRAY(ast_owner, asset_counter) = person
							Exit For
						End If
					Next
					ASSETS_ARRAY(ast_instance, asset_counter) = "0" & CARS_instance
					If CARS_type = "1" Then ASSETS_ARRAY(ast_type, asset_counter) = "1 - Car"
					If CARS_type = "2" Then ASSETS_ARRAY(ast_type, asset_counter) = "2 - Truck"
					If CARS_type = "3" Then ASSETS_ARRAY(ast_type, asset_counter) = "3 - Van"
					If CARS_type = "4" Then ASSETS_ARRAY(ast_type, asset_counter) = "4 - Camper"
					If CARS_type = "5" Then ASSETS_ARRAY(ast_type, asset_counter) = "5 - Motorcycle"
					If CARS_type = "6" Then ASSETS_ARRAY(ast_type, asset_counter) = "6 - Trailer"
					If CARS_type = "7" Then ASSETS_ARRAY(ast_type, asset_counter) = "7 - Other"
					ASSETS_ARRAY(ast_year, asset_counter) = CARS_year
					ASSETS_ARRAY(ast_make, asset_counter) = replace(CARS_make, "_", "")
					ASSETS_ARRAY(ast_model, asset_counter) = replace(CARS_model, "_", "")
					ASSETS_ARRAY(ast_trd_in, asset_counter) = trim(CARS_trade_in)
					ASSETS_ARRAY(ast_loan_value, asset_counter) = trim(CARS_loan)
					If CARS_source = "1" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "1 - NADA"
					If CARS_source = "2" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "2 - Appraisal Val"
					If CARS_source = "3" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "3 - Client Stmt"
					If CARS_source = "4" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "4 - Other Document"
					If CARS_owner_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1 - Title"
					If CARS_owner_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - License Reg"
					If CARS_owner_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - DMV"
					If CARS_owner_verif = "4" Then ASSETS_ARRAY(ast_verif, asset_counter) = "4 - Purchase Agmt"
					If CARS_owner_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
					If CARS_owner_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prvd"
					ASSETS_ARRAY(ast_amt_owed, asset_counter) = trim(replace(CARS_owe_amount, "_", ""))
					ASSETS_ARRAY(ast_owe_YN, asset_counter) = replace(CARS_joint_owner_YN, "_", "")
					ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(CARS_owed_date, " ", "/")
					If ASSETS_ARRAY(ast_bal_date, asset_counter) = "__/__/__" Then ASSETS_ARRAY(ast_bal_date, asset_counter) = ""
					If CARS_use = "1" Then ASSETS_ARRAY(ast_use, asset_counter) = "1 -  Primary Veh"
					If CARS_use = "2" Then ASSETS_ARRAY(ast_use, asset_counter) = "2 - Emp/Trng Trans/Emp Search"
					If CARS_use = "3" Then ASSETS_ARRAY(ast_use, asset_counter) = "3 - Disa Trans"
					If CARS_use = "4" Then ASSETS_ARRAY(ast_use, asset_counter) = "4 - Inc Producing"
					If CARS_use = "5" Then ASSETS_ARRAY(ast_use, asset_counter) = "5 - Used As Home"
					If CARS_use = "7" Then ASSETS_ARRAY(ast_use, asset_counter) = "7 - Unlicensed"
					If CARS_use = "8" Then ASSETS_ARRAY(ast_use, asset_counter) = "8 - Othr Countable"
					If CARS_use = "9" Then ASSETS_ARRAY(ast_use, asset_counter) = "9 - Unavailable"
					If CARS_use = "0" Then ASSETS_ARRAY(ast_use, asset_counter) = "0 - Long Distance Emp Travel"
					If CARS_use = "A" Then ASSETS_ARRAY(ast_use, asset_counter) = "A - Carry Heating Fuel Or Water"
					ASSETS_ARRAY(ast_hc_benefit, asset_counter) = CARS_hc_benefit
					ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = CARS_joint_owner_YN
					ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(CARS_share_ratio, " ", "")
					ASSETS_ARRAY(update_date, asset_counter) = replace(CARS_updated_date, " ", "/")
					ASSETS_ARRAY(update_panel, asset_counter) = unchecked

					transmit
					asset_counter = asset_counter + 1
					EMReadScreen reached_last_CARS_panel, 13, 24, 2
				Loop until reached_last_CARS_panel = "ENTER A VALID"
			End If
		Next

		current_asset_panel = FALSE
		acct_panels = 0
		secu_panels = 0
		cars_panels = 0
		For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
			If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
				current_asset_panel = TRUE
				acct_panels = acct_panels + 1
				If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
				asset_display = asset_display & vbNewLine & "ACCT - " & the_asset
			ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
				current_asset_panel = TRUE
				secu_panels = secu_panels + 1
				If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
				asset_display = asset_display & vbNewLine & "SECU - " & the_asset
			ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
				current_asset_panel = TRUE
				cars_panels = cars_panels + 1
				If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
				asset_display = asset_display & vbNewLine & "CARS - " & the_asset
			Else
				asset_display = asset_display & vbNewLine & ASSETS_ARRAY(ast_panel, the_asset) & " - " & the_asset
			End If
		Next
	End If
	'ASSET CODE-END
	If form_type_array(form_type_const, maxis_panel_read) = arep_form_name Then  'MAXIS NAVIGATION FOR AREP: READ AREP 
		Do
			Call navigate_to_MAXIS_screen("STAT", "AREP")
			EMReadScreen nav_check, 4, 2, 53
			EMWaitReady 0, 0
		Loop until nav_check = "AREP"

		arep_update_AREP_panel_checkbox = checked

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
		End If
	End If

	If form_type_array(form_type_const, maxis_panel_read) = hosp_form_name Then	'MAXIS NAVIGATION FOR HOSPICE: Seach casenotes back 1 year to find last HOSP Form Recieved, checks date of death 
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

		If hosp_exit_date <> "" Then     'if there is an exit date in the note found then we don't want to use the information from that note
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

	If form_type_array(form_type_const, maxis_panel_read) = diet_form_name Then	'MAXIS NAVIGATION FOR DIET: CASE CURR: Reading status of programs
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

	If form_type_array(form_type_const, maxis_panel_read) = psn_form_name Then	'MAXIS NAVIGATION FOR PSN: READ WREG, DISA
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
		EMReadScreen psn_disa_cert_start, 10, 7, 47
		EMReadScreen psn_disa_cert_end, 10, 7, 69
		EMReadScreen psn_disa_status, 2 , 11, 59
		EMReadScreen psn_disa_verif, 1, 11, 69
		psn_disa_begin_date = replace(psn_disa_begin_date,  "__ __ ____", "")
		psn_disa_end_date = replace(psn_disa_end_date,  "__ __ ____", "")
		psn_disa_begin_date = replace(psn_disa_begin_date, " ", "/")
		psn_disa_end_date = replace(psn_disa_end_date, " ", "/")
		psn_disa_cert_start = replace(psn_disa_cert_start,  "__ __ ____", "")
		psn_disa_cert_end = replace(psn_disa_cert_end,  "__ __ ____", "")
		psn_disa_cert_start = replace(psn_disa_cert_start, " ", "/")
		psn_disa_cert_end = replace(psn_disa_cert_end, " ", "/")
		psn_disa_status = replace(psn_disa_status, "__", "")
		psn_disa_verif = replace(psn_disa_verif, "_", "")

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

	'Only reading vendor numer and name from current faci panel 
		Do
			Call navigate_to_MAXIS_screen("STAT", "FACI")		'Navigate to FACI 
			EMReadScreen nav_check, 4, 2, 44
			EMWaitReady 0, 0
		Loop until nav_check = "FACI"
		EMWriteScreen Left(psn_member_dropdown, 2), 20, 76
		Call write_value_and_transmit("01", 20, 79)		'Make sure we are on page 01 for the member to start
			EMReadScreen existance_check, 1, 2, 73
			stat_faci_exists = True		
			If existance_check = "0" Then stat_faci_exists = False
			If stat_faci_exists = True Then
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
						If in_year_check_01 <> "____" and out_year_check_01 = "____" Then faci_row = 14
						If in_year_check_02 <> "____" and out_year_check_02 = "____" Then faci_row = 15
						If in_year_check_03 <> "____" and out_year_check_03 = "____" Then faci_row = 16
						If in_year_check_04 <> "____" and out_year_check_04 = "____" Then faci_row = 17
						If in_year_check_05 <> "____" and out_year_check_05 = "____" Then faci_row = 18
						exit do
					Elseif FACI_current_panel = FACI_total_check then
						currently_in_FACI = False
						exit do
					Else
						transmit
					End if
				Loop until FACI_current_panel = FACI_total_check
				'stat_faci_currently_in_facility = currently_in_FACI

				If currently_in_FACI = True then
					EmReadscreen psn_faci_vendor_number, 8, 5, 43
					EMReadScreen psn_faci_vendor_name, 30, 6, 43
					psn_faci_vendor_name = trim(replace(psn_faci_vendor_name, "_", ""))
					psn_faci_vendor_number = trim(replace(psn_faci_vendor_number, "_", ""))
					psn_facility = psn_faci_vendor_number & "-" & psn_faci_vendor_name
				End if
			End If

	End If

	If form_type_array(form_type_const, maxis_panel_read) = sf_form_name Then	'MAXIS NAVIGATION FOR PSN: READ MEMB, ADDR, HEST, SHEL
		'SEARCH THE LIST OF HOUSEHOLD MEMBERS TO SEARCH ALL SHEL PANELS
		Do 
			CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
			EMReadScreen nav_check, 4, 2, 48
			EMWaitReady 0, 0
		Loop until nav_check = "MEMB"	

		DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
			EMReadscreen ref_nbr, 3, 4, 33
			EMReadScreen access_denied_check, 13, 24, 2
			If access_denied_check = "ACCESS DENIED" Then
				PF10
				EMWaitReady 0, 0
				last_name = "UNABLE TO FIND"
				first_name = " - Access Denied"
				mid_initial = ""
			Else
				client_array = client_array & ref_nbr & "|"
			End If
			transmit
			Emreadscreen edit_check, 7, 24, 2
		LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

		client_array = TRIM(client_array)
		If right(client_array, 1) = "|" Then client_array = left(client_array, len(client_array) - 1)
		ref_numbers_array = split(client_array, "|")

		members_counter = 0
		btn_placeholder = 600
		member_selection = ""
		For each memb_ref_number in ref_numbers_array
			Do
				Call navigate_to_MAXIS_screen("STAT", "SHEL")
				EMReadScreen nav_check, 4, 2, 48
				EmWaitReady 0, 0
			Loop until nav_check = "SHEL"
			EMWriteScreen memb_ref_number, 20, 76
			transmit

			ReDim Preserve ALL_SHEL_PANELS_ARRAY(shel_entered_notes_const, members_counter)
			ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, members_counter) = memb_ref_number
			ALL_SHEL_PANELS_ARRAY(memb_btn_const, members_counter) = btn_placeholder + members_counter
			ALL_SHEL_PANELS_ARRAY(attempted_update_const, members_counter) = False

			EMReadScreen shel_version, 1, 2, 73
			If shel_version = "1" Then
				ALL_SHEL_PANELS_ARRAY(shel_exists_const, members_counter) = True
				If total_paid_to = "" Then total_paid_to =  ALL_SHEL_PANELS_ARRAY(paid_to_const, members_counter)
				' If member_selection = "" Then member_selection = members_counter
			Else
				ALL_SHEL_PANELS_ARRAY(shel_exists_const, members_counter) = False
				ALL_SHEL_PANELS_ARRAY(original_panel_info_const, members_counter) = "||||||||||||||||||||||||||||||||||"
			End If

			Do
				Call navigate_to_MAXIS_screen("STAT", "MEMB")
				EMReadScreen nav_check, 4, 2, 48
				EmWaitReady 0, 0
			Loop until nav_check = "MEMB"
			EMWriteScreen memb_ref_number, 20, 76
			transmit
			EMReadScreen memb_panel_age, 3, 8, 76
			memb_panel_age = trim(memb_panel_age)
			If memb_panel_age = "" Then memb_panel_age = 0
			memb_panel_age = memb_panel_age * 1
			ALL_SHEL_PANELS_ARRAY(person_age_const, members_counter) = memb_panel_age

			If ALL_SHEL_PANELS_ARRAY(shel_exists_const, members_counter) = True Then ALL_SHEL_PANELS_ARRAY(person_shel_checkbox, members_counter) = checked
			members_counter = members_counter + 1
		Next

		Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
		Call access_HEST_panel("READ", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
		For shel_member = 0 to UBound(ALL_SHEL_PANELS_ARRAY, 2)
			If ALL_SHEL_PANELS_ARRAY(shel_exists_const, shel_member) = True Then
				Call access_SHEL_panel("READ", ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, shel_member), ALL_SHEL_PANELS_ARRAY(hud_sub_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(shared_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(paid_to_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member))
			End If
		Next
		page_to_display = ADDR_dlg_page
		Call read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, total_shelter_expense, total_shel_original_information)
		
		'here we save the information we gathered to start with so that we can compare it and know if it changed
		hest_original_information = all_persons_paying&"|"&all_persons_paying&"|"&choice_date&"|"&actual_initial_exp&"|"&retro_heat_ac_yn&"|"&_
		retro_heat_ac_units&"|"&retro_heat_ac_amt&"|"&retro_electric_yn&"|"&retro_electric_units&"|"&retro_electric_amt&"|"&retro_phone_yn&"|"&_
		retro_phone_units&"|"&retro_phone_amt&"|"&prosp_heat_ac_yn&"|"&prosp_heat_ac_units&"|"&prosp_heat_ac_amt&"|"&prosp_electric_yn&"|"&_
		prosp_electric_units&"|"&prosp_electric_amt&"|"&prosp_phone_yn&"|"&prosp_phone_units&"|"&prosp_phone_amt&"|"&total_utility_expense

		hest_original_information = UCASE(hest_original_information)
		addr_update_attempted = False
		shel_update_attempted = False
		hest_update_attempted = False
		resi_line_one = ""
		resi_line_two = ""
		mail_line_one = ""
		mail_line_two = ""
	End If
Next

'DIALOG DISPLAYING FORM SPECIFIC INFORMATION===========================================================================
'Displays individual dialogs for each form selected via checkbox or dropdown. Do/Loops allows us to jump around/are more flexible than For/Next 
form_count = 0

Do
	Do
		Do
			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 456, 300, "Documents Received - Case #" & MAXIS_case_number
				If form_type_array(form_type_const, form_count) = asset_form_name then 
					Call asset_dialog
					current_dialog = "asset"
				End If
				If form_type_array(form_type_const, form_count) = atr_form_name Then 
					Call atr_dialog
					current_dialog = "atr"
				End If
				If form_type_array(form_type_const, form_count) = arep_form_name then 
					Call arep_dialog
					current_dialog = "arep"
				End If
				If form_type_array(form_type_const, form_count) = change_form_name Then 
					Call change_dialog
					current_dialog = "chng"
				End If
				If form_type_array(form_type_const, form_count) = evf_form_name Then 
					Call evf_dialog
					current_dialog = "evf"
				End If
				If form_type_array(form_type_const, form_count) = hosp_form_name Then 
					Call hospice_dialog
					current_dialog = "hosp"
				End If
				If form_type_array(form_type_const, form_count) = iaa_form_name Then 
					Call iaa_dialog
					current_dialog = "iaa"
				End If
				If form_type_array(form_type_const, form_count) = ltc_1503_form_name Then 
					Call ltc_1503_dialog
					current_dialog = "ltc 1503"
				End If
				If form_type_array(form_type_const, form_count) = mof_form_name Then 
					Call mof_dialog
					current_dialog = "mof"
				End If
				If form_type_array(form_type_const, form_count) = mtaf_form_name Then 
					Call mtaf_dialog
					current_dialog = "mtaf"
				End If
				If form_type_array(form_type_const, form_count) = psn_form_name Then 
					Call psn_dialog
					current_dialog = "psn"
				End If
				If form_type_array(form_type_const, form_count) = sf_form_name Then 
					Call sf_dialog
					current_dialog = "sf"
				End If
				If form_type_array(form_type_const, form_count) = diet_form_name Then 
					Call diet_dialog
					current_dialog = "diet"
				End If
				If form_type_array(form_type_const, form_count) = other_form_name Then 
					Call other_dialog
					current_dialog = "other"
				End If
				
				If left(docs_rec, 2) = ", " Then docs_rec = right(docs_rec, len(docs_rec)-2)        'trimming the ',' off of the list of docs

				btn_pos = 45		'variable to iterate down for each necessary button
				For current_form = 0 to Ubound(form_type_array, 2) 		'This iterates through the array and creates buttons for each form selected from top down. Also stores button name and number in the array based on form name selected. 
					If form_type_array(form_type_const, current_form) = asset_form_name then
						form_type_array(btn_name_const, form_count) = "ASSET"
						form_type_array(btn_number_const, form_count) = 400
						PushButton 395, btn_pos, 45, 15, "ASSET", asset_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = atr_form_name Then 
						form_type_array(btn_name_const, form_count) = "ATR"
						form_type_array(btn_number_const, form_count) = 401
						PushButton 395, btn_pos, 45, 15, "ATR", atr_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = arep_form_name then 
						form_type_array(btn_name_const, form_count) = "AREP"
						form_type_array(btn_number_const, form_count) = 402
						PushButton 395, btn_pos, 45, 15, "AREP", arep_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = change_form_name  then 
						form_type_array(btn_name_const, form_count) = "CHNG"
						form_type_array(btn_number_const, form_count) = 403
						PushButton 395, btn_pos, 45, 15, "CHNG", change_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = evf_form_name  then 
						form_type_array(btn_name_const, form_count) = "EVF"
						form_type_array(btn_number_const, form_count) = 404		
						PushButton 395, btn_pos, 45, 15, "EVF", evf_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = hosp_form_name  then 
						form_type_array(btn_name_const, form_count) = "HOSP"
						form_type_array(btn_number_const, form_count) = 405
						PushButton 395, btn_pos, 45, 15, "HOSP", hospice_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = iaa_form_name  then 
						form_type_array(btn_name_const, form_count) = "IAA"
						form_type_array(btn_number_const, form_count) = 406
						PushButton 395, btn_pos, 45, 15, "IAA", iaa_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = ltc_1503_form_name then 
						form_type_array(btn_name_const, form_count) = "LTC-1503"
						form_type_array(btn_number_const, form_count) = 408
						PushButton 395, btn_pos, 45, 15, "LTC-1503", ltc_1503_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = mof_form_name then 
						form_type_array(btn_name_const, form_count) = "MOF"
						form_type_array(btn_number_const, form_count) = 409
						PushButton 395, btn_pos, 45, 15, "MOF", mof_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = mtaf_form_name then 
						form_type_array(btn_name_const, form_count) = "MTAF"
						form_type_array(btn_number_const, form_count) = 410
						PushButton 395, btn_pos, 45, 15, "MTAF", mtaf_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = psn_form_name then 
						form_type_array(btn_name_const, form_count) = "PSN"
						form_type_array(btn_number_const, form_count) = 411
						PushButton 395, btn_pos, 45, 15, "PSN", psn_btn 
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = sf_form_name then 
						form_type_array(btn_name_const, form_count) = "SF"
						form_type_array(btn_number_const, form_count) = 412
						PushButton 395, btn_pos, 45, 15, "SF", sf_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = diet_form_name then 
						form_type_array(btn_name_const, form_count) = "DIET"
						form_type_array(btn_number_const, form_count) = 413
						PushButton 395, btn_pos, 45, 15, "DIET", diet_btn
						btn_pos = btn_pos + 15
					End If
					If form_type_array(form_type_const, current_form) = other_form_name then 
						form_type_array(btn_name_const, form_count) = "OTHR"
						form_type_array(btn_number_const, form_count) = 414
						PushButton 395, btn_pos, 45, 15, "OTHR", other_btn
						btn_pos = btn_pos + 15
					End If
					'MsgBox "Current form" & form_type_array(form_type_const, current_form)
				Next

				Text 395, 35, 45, 10, "    --Forms--"
				If form_count > 0 Then PushButton 395, 255, 50, 15, "Previous", previous_btn ' Previous button to navigate from one form to the previous one.
				If form_count < Ubound(form_type_array, 2) Then PushButton 395, 275, 50, 15, "Next Form", next_btn	'Next button to navigate from one form to the next. 
				If form_count = Ubound(form_type_array, 2) Then PushButton 395, 275, 50, 15, "Complete", complete_btn	'Complete button kicks off the casenoting of all completed forms. 
				'MsgBox "Ubound(form_type_array, 2)" & Ubound(form_type_array, 2) 			
			EndDialog

			err_msg = ""
			asset_err_msg = ""
			atr_err_msg = ""
			arep_err_msg = ""
			chng_err_msg = ""
			evf_err_msg = "" 
			hosp_err_msg = ""
			iaa_err_msg = ""
			ltc_1503_err_msg = ""
			mof_err_msg = ""
			mtaf_err_msg = ""
			psn_err_msg = ""
			sf_err_msg = ""
			diet_err_msg = ""
			other_err_msg = "" 

			dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
			cancel_confirmation
			
			Call form_specific_error_handling	'function for error handling of main dialog of forms 
			Call dialog_movement				'function to move throughout the dialogs
		Loop until err_msg = ""
	Loop until ButtonPressed = complete_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'DIALOG OUTSTANDING VERIFICATIONS===========================================================================
Do
	DO
		err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 321, 50, "Outstanding Verifications"
			EditBox 120, 5, 195, 15, outstanding_verifs
			ButtonGroup ButtonPressed
				OkButton 210, 30, 50, 15
				CancelButton 265, 30, 50, 15
			Text 5, 10, 115, 10, "Specify outstanding verifications:"
		EndDialog

		dialog Dialog1	'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'WRITE IN MAXIS===========================================================================
Call MAXIS_background_check
For maxis_panel_write = 0 to Ubound(form_type_array, 2)	

	If form_type_array(form_type_const, maxis_panel_write) = arep_form_name then 		' WRITE FOR AREP: Updates AREP panel from either AREP_recvd_date or arep_signature_date through CM+1
		'formatting programs into one variable to write in case note
		IF arep_SNAP_AREP_checkbox = checked THEN AREP_programs = "SNAP"
		IF arep_HC_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", HC"
		IF arep_CASH_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", CASH"
		If left(AREP_programs, 1) = "," Then AREP_programs = right(AREP_programs, len(AREP_programs)-2)

		If arep_update_AREP_panel_checkbox = checked Then				'If update AREP checkbox is checked, then update panel 	
			If IsDate(arep_signature_date) = TRUE Then 
				Call convert_date_into_MAXIS_footer_month(arep_signature_date, MAXIS_footer_month, MAXIS_footer_year)
			Else		 
				Call convert_date_into_MAXIS_footer_month(AREP_recvd_date, MAXIS_footer_month, MAXIS_footer_year)
			End If
				
			Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)
			For each item in date_array
				MAXIS_footer_month = datepart("m", item)
				If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
				MAXIS_footer_year = right(datepart("yyyy", item), 2)
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
			Next	
		End If
    End If

	If form_type_array(form_type_const, maxis_panel_write) = iaa_form_name Then			'WRITE FOR IAA: If iaa_update_pben_checkbox = checked updates PBEN from iaa_date_received through CM+1. If all rows are full, script will stop user. 
		If iaa_update_pben_checkbox = checked Then 
			pben_updated = FALSE
			iaa_referral_date_month = right("00" & DatePart("m", iaa_referral_date), 2)		'Setting up the parts of the date for MAXIS fields
			iaa_referral_date_day = right("00" & DatePart("d", iaa_referral_date), 2)
			iaa_referral_date_year = right(DatePart("yyyy", iaa_referral_date), 2)
			
			If iaa_date_applied_pben <> "" Then 
				iaa_date_applied_pben_month = right("00" & DatePart("m", iaa_date_applied_pben), 2)		'Setting up the parts of the date for MaXIS fields	
				iaa_date_applied_pben_day = right("00" & DatePart("d", iaa_date_applied_pben), 2)
				iaa_date_applied_pben_year = right(DatePart("yyyy", iaa_date_applied_pben), 2)
			End If
			
			If iaa_iaa_date <> "" Then 
				iaa_date_month = right("00" & DatePart("m", iaa_iaa_date), 2)		'Setting up the parts of the date for MAXIS fields	
				iaa_date_day = right("00" & DatePart("d", iaa_iaa_date), 2)
				iaa_date_year = right(DatePart("yyyy", iaa_iaa_date), 2)
				pben_member_number = Left(iaa_member_dropdown, 2)
			End If 

			Call convert_date_into_MAXIS_footer_month(iaa_date_received, MAXIS_footer_month, MAXIS_footer_year)
			Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)
			For each item in date_array
				MAXIS_footer_month = datepart("m", item)
				If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
				MAXIS_footer_year = right(datepart("yyyy", item), 2)
				Do
					Call Navigate_to_MAXIS_screen ("STAT", "PBEN")					'Go to PBEN 
					EMReadScreen nav_check, 4, 2, 49
					EMWaitReady 0, 0
				Loop until nav_check = "PBEN"
								
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
						pben_updated = TRUE									
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
						iaa_update_pben_checkbox = checked
						Exit Do

					ElseIf pben_exist = "02" Then 
						If Left(iaa_benefit_type, 2) = "02" Then 								'If 02 benefit type already exists, must evaluate to see if it is AEPN status. If so, we cannot update the panel. 
							EMReadScreen pben_benefit_type, 2, pben_row, 24
							EMReadScreen pben_referral_date, 8, pben_row, 40
							EMReadScreen pben_date_applied, 8, pben_row, 51
							EMReadScreen pben_verification, 1, pben_row, 62
							EMReadScreen pben_iaa_date, 8, pben_row, 66
							EMReadScreen pben_disp_code, 1, pben_row, 77
							pben_disp_code_string = pben_disp_code_string & pben_disp_code		
					
							If Instr(pben_disp_code_string, "A") or Instr(pben_disp_code_string, "E") or Instr(pben_disp_code_string, "P") or Instr(pben_disp_code_string, "N") Then 
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
			Next
		End If
	End If 

	If form_type_array(form_type_const, maxis_panel_write) = ltc_1503_form_name then 	' WRITE FOR LTC 1503:  Verifies Max number of FACI panels have not been met for ltc_1503_faci_footer_month through CM+1. Then updates FACI panel from ltc_1503_faci_footer_month through CM+1. 
		end_msg = end_msg & vbNewLine & "LTC 1503 Form information entered."
		Original_footer_month = MAXIS_footer_month
		Original_footer_year = MAXIS_footer_year
		MAXIS_footer_month = ltc_1503_faci_footer_month
		MAXIS_footer_year = ltc_1503_faci_footer_year
		If MAXIS_footer_month = "" AND MAXIS_footer_year = "" Then 
			MAXIS_footer_month = Original_footer_month
			MAXIS_footer_year = Original_footer_year
		End If 
		Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)

		If ltc_1503_FACI_update_checkbox = checked then		'If update FACI checkbox checked udpate panel
			For each item in date_array
				MAXIS_footer_month = datepart("m", item)
				If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
				MAXIS_footer_year = right(datepart("yyyy", item), 2)
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
					ltc_1503_FACI_update_checkbox = checked
				END IF
			Next 

			If ltc_1503_FACI_update_checkbox = checked then		'If update FACI checkbox checked udpate panel
				ltc_1503_updated_FACI_checkbox = checked
				For each item in date_array
					MAXIS_footer_month = datepart("m", item)
					If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
					MAXIS_footer_year = right(datepart("yyyy", item), 2)
					Do
						call navigate_to_MAXIS_screen("STAT", "FACI")	'Navigate to FACI
						EMReadScreen nav_check, 4, 2, 44
						EMWaitReady 0, 0
					Loop until nav_check = "FACI"
					EMWriteScreen "NN", 20, 79
					transmit
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
				Next 
			End If
		End if

		'HCMI
		If ltc_1503_HCMI_update_checkbox = checked THEN
			For each item in date_array
				MAXIS_footer_month = datepart("m", item)
				If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
				MAXIS_footer_year = right(datepart("yyyy", item), 2)
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
			Next
		END IF
		MAXIS_footer_month = Original_footer_month
		MAXIS_footer_year = Original_footer_year
	End If

	If form_type_array(form_type_const, maxis_panel_write) = mtaf_form_name then 		'MANUAL WRITE FOR MTAF: Promps user to update PROG if it does not meet requirements
		Do
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			EMReadScreen nav_check, 4, 2, 50
			EmWaitReady 0, 0
		Loop until nav_check = "PROG"
		EMReadScreen prog_cash_1_status, 4, 6, 74
			If prog_cash_1_status = "PEND" Then
				EMReadScreen prog_cash_1_intvw_date, 8, 6, 55
				prog_cash_1_intvw_date = replace(prog_cash_1_intvw_date, " ", "/")
				If prog_cash_1_intvw_date = "__/__/__" Then prog_cash_1_intvw_date = ""
				If prog_cash_1_intvw_date = "" Then update_prog = True

			End If
			EMReadScreen prog_cash_2_status, 4, 7, 74
			If prog_cash_2_status = "PEND" Then
				EMReadScreen prog_cash_2_intvw_date, 8, 7, 55
				prog_cash_2_intvw_date = replace(prog_cash_2_intvw_date, " ", "/")
				If prog_cash_2_intvw_date = "__/__/__" Then prog_cash_2_intvw_date = ""
				If prog_cash_2_intvw_date = "" Then update_prog = True
			End If
			If update_prog = True Then
				Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 251, 140, "Update Interview Date in STAT"
					ButtonGroup ButtonPressed
						OkButton 195, 120, 50, 15
					Text 30, 10, 200, 10, "It appears that PROG is not updated with an Interview Date."
					GroupBox 10, 30, 230, 45, "UPDATE PROG NOW"
					Text 30, 50, 200, 10, "Update PROG with and Interview Date for PENDING CASH."
					Text 10, 85, 230, 35, "To prevent unnecessary notices, we code the interview date for any pending program that does not require an interview. match the Interview Date to the Application Date for the CASH program pending with no interview date."
					Text 10, 125, 115, 10, "Press OK when PROG is Updated."
				EndDialog

				dialog Dialog1	'Calling a dialog without a assigned variable will call the most recently defined dialog
			End If
	End If
		
	If form_type_array(form_type_const, maxis_panel_write) = psn_form_name then 		'WRITE FOR PSN: Updates DISA and WREG from psn_date_received through CM+1
		If psn_udpate_wreg_disa_checkbox = checked Then 
			Call convert_date_into_MAXIS_footer_month(psn_date_received, MAXIS_footer_month, MAXIS_footer_year)
			Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)
			For each item in date_array
				MAXIS_footer_month = datepart("m", item)
				If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
				MAXIS_footer_year = right(datepart("yyyy", item), 2)
				Do																'Function to write information to DISA
					Call Navigate_to_MAXIS_screen ("STAT", "DISA")				'Goes to DISA for the correct person
					EMReadScreen nav_check, 4, 2, 45
				Loop until nav_check = "DISA"
				EMWriteScreen Left(psn_member_dropdown, 2), 20, 76
				transmit
				EMReadScreen exist_check, 1, 2, 73
					If exist_check = "0" Then 
						Call write_value_and_transmit ("NN", 20, 79)
					Else	
						PF9
					End If
					EmWaitReady 0, 0

					'Handling for fields left blank in the dialog			
					If psn_disa_begin_date = "" Then 
						disa_start_month = "__"
						disa_start_day = "__"
						disa_start_year = "____"
					Else 
						disa_start_month = right("00" & DatePart("m", psn_disa_begin_date), 2)	'Isolates the start month, day, and year as these are seperate fields on DISA
						disa_start_day = right("00" & DatePart("d", psn_disa_begin_date), 2)
						disa_start_year = DatePart("yyyy", psn_disa_begin_date)
					End If
					
					If psn_disa_end_date = "" Then 
						disa_end_month = "__"
						disa_end_day = "__"
						disa_end_year = "____"
					Else 
						disa_end_month = right("00" & DatePart("m", psn_disa_end_date), 2)		'Isolates the end month, day, and year as these are seperate fields on DISA
						disa_end_day = right("00" & DatePart("d", psn_disa_end_date), 2)
						disa_end_year = DatePart("yyyy", psn_disa_end_date)
					End If

					If psn_disa_cert_start = "" Then 
						cert_start_month = "__"
						cert_start_day = "__"
						cert_start_year = "____"
					Else 
						cert_start_month = right("00" & DatePart("m", psn_disa_cert_start), 2)	'Isolates the start month, day, and year as these are seperate fields on DISA
						cert_start_day = right("00" & DatePart("d", psn_disa_cert_start), 2)
						cert_start_year = DatePart("yyyy", psn_disa_cert_start)
					End If 

					If psn_disa_cert_end = "" Then 
						cert_end_month = "__"
						cert_end_day = "__"
						cert_end_year = "____"
					Else 
						cert_end_month = right("00" & DatePart("m", psn_disa_cert_end), 2)		'Isolates the end month, day, and year as these are seperate fields on DISA
						cert_end_day = right("00" & DatePart("d", psn_disa_cert_end), 2)
						cert_end_year = DatePart("yyyy", psn_disa_cert_end)
					End If

					'Writing the Disability Begin Date'
					EMWriteScreen disa_start_month, 6, 47
					EMWriteScreen disa_start_day, 6, 50
					EMWriteScreen disa_start_year, 6, 53
					
					'Writing the Disability End Date'
					EMWriteScreen disa_end_month, 6, 69
					EMWriteScreen disa_end_day, 6, 72
					EMWriteScreen disa_end_year, 6, 75
					
					'Writing the Certification Begin Date'
					EMWriteScreen cert_start_month, 7, 47
					EMWriteScreen cert_start_day, 7, 50
					EMWriteScreen cert_start_year, 7, 53
					
					'Writing the Certification End Date'
					EMWriteScreen cert_end_month, 7, 69
					EMWriteScreen cert_end_day, 7, 72
					EMWriteScreen cert_end_year, 7, 75

					'Writing the disa status and verif code'
					EMWriteScreen psn_disa_status, 11, 59
					EMWriteScreen psn_disa_verif, 11, 69
					transmit
				Do																'Function to write information to DISA
					Call Navigate_to_MAXIS_screen("STAT", "WREG")				'Goes to DISA for the correct person
					EMReadScreen nav_check, 4, 2, 48
					EMWaitReady 0, 0
				Loop until nav_check = "WREG"
				EMWriteScreen Left(psn_member_dropdown, 2), 20, 76
				transmit
				EMReadScreen exist_check, 1, 2, 73
				If exist_check = "0" Then 
					Call write_value_and_transmit ("NN", 20, 79)
				Else	
					PF9
				End If
				EMWriteScreen Left(psn_wreg_fs_pwe, 1), 6, 68						
				EMWriteScreen Left(psn_wreg_work_wreg_status, 2), 8, 50			
				EMWriteScreen Left(psn_wreg_abawd_status, 2), 13, 50				
				EMWriteScreen Left(psn_wreg_ga_elig_status, 2),  15, 50			
				transmit	
			Next
		End If
	End If

	If form_type_array(form_type_const, maxis_panel_write) = diet_form_name Then		'Write FOR DIET: Only updates DIET Panel from diet_date_received through CM+1 if diet is approved and active/pending MSA/MFIP. 

		If diet_status_dropdown = "Approved" Then			'Only if the diet is approved should we update the pben panel
			If diet_mfip_msa_status <> "Not Active/Pending" Then		'Only if the determine program case status determines the case is active or pending on MSA or MFIP will it fill out the DIET panel. 
				Call convert_date_into_MAXIS_footer_month(diet_date_received, MAXIS_footer_month, MAXIS_footer_year)
				Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)
				For each item in date_array
					MAXIS_footer_month = datepart("m", item)
					If len(MAXIS_footer_month) = 1 Then MAXIS_footer_month = "0" & MAXIS_footer_month 
					MAXIS_footer_year = right(datepart("yyyy", item), 2)
					updated_diet_months = updated_diet_months & MAXIS_footer_month & "/" & MAXIS_footer_year & "," 
					Do
						Call navigate_to_MAXIS_screen("STAT", "DIET")
						EMReadScreen nav_check, 4, 2, 48
						EmWaitReady 0, 0
					Loop until nav_check = "DIET"
					diet_ref_number = Left(diet_member_number, 2)					'Grabbing member number from the member dropdown selection
					Call write_value_and_transmit(diet_ref_number, 20, 76)			'Go to the correct member 
					EMReadScreen DIET_total, 1, 2, 78
					If DIET_total = 0 then 								'If panel count is 0, then create a panel
						Call write_value_and_transmit("NN", 20, 79)		
					Else												'If panel exists, edit mode, delete panel, create new panel
						PF9
						EMWaitReady 0, 0
						Call write_value_and_transmit("DEL", 20, 71)	
						EMWaitReady 0, 0
						Call write_value_and_transmit("NN", 20, 79)	
						EMWaitReady 0, 0
					End If
					If diet_mfip_msa_status = "MFIP-Active" or diet_mfip_msa_status = "MFIP-Pending" Then		'If MFIP then write in diet, hard coded
						EMWriteScreen left(diet_1_dropdown, 2), 8, 40
						EMWriteScreen left(diet_verif_1_dropdown, 1), 8, 51
						EMWriteScreen left(diet_2_dropdown, 2), 9, 40
						EMWriteScreen left(diet_verif_2_dropdown, 1), 9, 51	
						transmit 	
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
						transmit
					End If
				Next
				If right(updated_diet_months, 1) = "," Then updated_diet_months = left(updated_diet_months, len(updated_diet_months)-1)
			End If
		End If
	End if
Next

'TIKLS===========================================================================
If AREP_TIKL_check = checked then 'AREP TIKLS
	Call create_TIKL("Client's AREP release for HC is now 12 months old and no longer valid. Take appropriate action.", 365, arep_signature_date, False, TIKL_note_text)
	end_msg = end_msg & vbNewLine & "AREP: TIKL has been sent for a year from now to request an updated form."
End IF

If EVF_TIKL_checkbox = checked Then 'EVF TIKL
	Call create_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action.", 10, date, True, TIKL_note_text)
	end_msg = end_msg & vbNewLine & "EVF: TIKL has been sent for 10 days from now for the additional information requested."
End If

If ltc_1503_TIKL_checkbox = checked Then 'LTC 1503 TIKL
	If ltc_1503_length_of_stay = "30 days or less"   then ltc_1503_TIKL_multiplier = 30
	If ltc_1503_length_of_stay = "31 to 90 days"     then ltc_1503_TIKL_multiplier = 90
	If ltc_1503_length_of_stay = "91 to 180 days"    then ltc_1503_TIKL_multiplier = 180
	Call create_TIKL("Have " & worker_signature & " call " & ltc_1503_FACI_1503 & " re: length of stay. " & ltc_1503_TIKL_multiplier & " days expired.", ltc_1503_TIKL_multiplier, ltc_1503_admit_date, False, TIKL_note_text)
	end_msg = end_msg & vbNewLine & "LTC1503: TIKL has been sent for " & ltc_1503_TIKL_multiplier & " days from now for the additional information requested."
End If

If psn_tikl_checkbox = checked then 'PSN
	Call create_TIKL("Resident's PSN is now 12 months old and no longer valid. Take appropriate action.", 365, psn_date_received, False, TIKL_note_text)
	end_msg = end_msg & vbNewLine & "PSN: TIKL has been sent for a year from now to request an updated form."
End If

If sf_tikl_nav_check = 1 then 'SF TIKL
	Call create_TIKL("CHANGE REPORTED", 365, date, False, TIKL_note_text)
	end_msg = end_msg & vbNewLine & "Shelter: TIKL has been sent for a year from now to request an updated form."
End If 

If diet_tikl_checkbox = 1 Then 'DIET TIKL
	Call create_TIKL("Special Diet approaching renewal", 305, diet_date_received, False, TIKL_note_text) 
	end_msg = end_msg & vbNewLine & "DIET: TIKL has been sent for 10 mo from now for renewal."
End If

'CREATING LIST OF DOCS_REC AND END_MSG=========================================================================== 
'Documents case noted as their own unique casenote do not have handling for docs_rec
For list_of_docs_received = 0 to Ubound(form_type_array, 2) 
	If form_type_array(form_type_const, list_of_docs_received) = asset_form_name then 
		If InStr(docs_rec,"ASST") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", ASST"
		End If
		If InStr(end_msg, "Asset detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "Asset detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = atr_form_name Then 
		If InStr(end_msg, "ATR detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "ATR detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = arep_form_name then 
		If InStr(end_msg, "AREP detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "AREP detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = change_form_name Then 
		If InStr(docs_rec,"CHNG") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", CHNG"
		End If
		If InStr(end_msg, "Change detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "Change detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = evf_form_name Then 
		If InStr(docs_rec,"EVF") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", EVF"
		End If
		If InStr(end_msg, "EVF detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "EVF detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = hosp_form_name Then 
		If InStr(end_msg, "Hospice detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "Hospice detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = iaa_form_name Then 
		If InStr(docs_rec,"IAA(s)") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", IAA(s)"
		End If
		If InStr(end_msg,"IAA(s) detail entered") Then 
			end_msg = end_msg
		Else
			end_msg = end_msg & vbNewLine & "IAA(s) detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = ltc_1503_form_name Then 
		If InStr(end_msg,"LTC-1503 detail entered") Then 
			end_msg = end_msg
		Else
			end_msg = end_msg & vbNewLine & "LTC-1503 detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = mof_form_name Then 
		If InStr(docs_rec,"MOF") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", MOF"
		End If
		If InStr(end_msg, "MOF detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "MOF detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = mtaf_form_name Then 
		If InStr(docs_rec,"MTAF") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", MTAF"
		End If
		If InStr(end_msg, "MTAF detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "MTAF detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = psn_form_name Then 
		If InStr(docs_rec,"PSN") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", PSN"
		End If
		If InStr(end_msg, "PSN detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "PSN detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = sf_form_name Then 
		If InStr(docs_rec,"SF") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", SF"
		End If
		If InStr(end_msg, "Shelter detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "Shelter detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = diet_form_name Then 
		If InStr(docs_rec,"DIET") Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec & ", DIET"
		End If
		If InStr(end_msg, "DIET detail entered") Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "DIET detail entered"
		End If
	End If
	If form_type_array(form_type_const, list_of_docs_received) = other_form_name Then 
		If InStr(docs_rec, other_list_form_names) Then
			docs_rec = docs_rec
		Else
			docs_rec = docs_rec  & ", " & other_list_form_names
		End If
		If InStr(end_msg, "Other Forms: " & other_list_form_names) Then 
			end_msg = end_msg
		Else 
			end_msg = end_msg & vbNewLine & "Other Forms: " & other_list_form_names & " detail entered"
		End If
	End If
Next 
If left(docs_rec, 2) = ", " Then docs_rec = right(docs_rec, len(docs_rec)-2)        'trimming the ',' off of the list of docs

'CASE NOTE===========================================================================
'For/Next creates one casenote for all documents received that should be CASENOTED TOGETHER. 
For each_case_note = 0 to Ubound(form_type_array, 2)	
	'Handling to change the case note header depending on if MTAF is one of the documents processed 
	If each_case_note = 0 Then
		If Instr(docs_rec, "MTAF") AND MTAF_note_only_checkbox = checked Then 
			Call start_a_blank_case_note
			CALL write_variable_in_CASE_NOTE("*** MTAF Processed: " & MTAF_status_dropdown & "***")
		ElseIf form_type_array(form_type_const, each_case_note) = mtaf_form_name OR form_type_array(form_type_const, each_case_note) = asset_form_name OR form_type_array(form_type_const, each_case_note) = change_form_name OR form_type_array(form_type_const, each_case_note) = evf_form_name OR form_type_array(form_type_const, each_case_note) = iaa_form_name OR form_type_array(form_type_const, each_case_note) = mof_form_name OR form_type_array(form_type_const, each_case_note) = psn_form_name OR form_type_array(form_type_const, each_case_note) = sf_form_name OR form_type_array(form_type_const, each_case_note) = diet_form_name OR form_type_array(form_type_const, each_case_note) = other_form_name Then 
			Call start_a_blank_case_note
			Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
		Else 
			Exit For
		End If
	End If

	If form_type_array(form_type_const, each_case_note) = asset_form_name then 		'Asset Statement Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** ASSET STATEMENT RECEIVED ***")
		CALL write_bullet_and_variable_in_CASE_NOTE("Date received", asset_date_received)
			If asset_dhs_6054_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("* Signed Personal Statement about Assets for Cash Received (DHS 6054)")
                Call write_bullet_and_variable_in_CASE_NOTE("  - Received on", st_verif_date)
                If signed_by_one <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_one & " on: " & signed_one_date)
                If signed_by_two <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_two & " on: " & signed_two_date)
                If signed_by_three <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_three & " on: " & signed_three_date)
                If box_one_info <> "" Then Call write_variable_in_CASE_NOTE("  - Account detail from form: " & box_one_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                        Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
                    End If
                Next
                If box_two_info <> "" Then Call write_variable_in_CASE_NOTE("  - Securities detail from form: " & box_two_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " Value: $" & ASSETS_ARRAY(ast_csv, the_asset) & " - Verif: " & left(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
                    End If
                Next
                If box_three_info <> "" Then Call write_variable_in_CASE_NOTE("  - Vehicle detail from form: " & box_three_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & ASSETS_ARRAY(ast_owe_verif, the_asset))
                    End If
                Next
            End If

			If asset_dhs_6054_checkbox = unchecked Then
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                        Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
                    End If
                Next

                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                        If left(ASSETS_ARRAY(ast_type, the_asset), 2) <> "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset))
                        If left(ASSETS_ARRAY(ast_type, the_asset), 2) = "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset) & " LI Face Value: $" & ASSETS_ARRAY(ast_face_value, the_asset))
                        Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                    End If
                Next

                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset))
                        Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_owe_verif, the_asset), len(ASSETS_ARRAY(ast_owe_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                    End If
                Next
            End If
			If Right(actions_taken, 2) = ", " Then actions_taken = left(actions_taken, len(actions_taken)-2)
			call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
		
		Call write_variable_in_case_note("---")
	End If

	If form_type_array(form_type_const, each_case_note) = change_form_name Then 	'Change Reported Case Note
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** CHANGE REPORT FORM RECEIVED ***")
		CALL write_bullet_and_variable_in_case_note("Notable changes reported", chng_notable_change)
		CALL write_bullet_and_variable_in_case_note("Effective Date", chng_effective_date)
		CALL write_bullet_and_variable_in_case_note("Date Received", chng_date_received)
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
	End If

	If form_type_array(form_type_const, each_case_note) = evf_form_name Then 		'EVF Case Notes
		Call start_a_blank_case_note
		Call write_variable_in_case_note("*** EVF FORM RECEIVED ***")
		Call write_bullet_and_variable_in_case_note("EVF received",  evf_date_received & "- " & EVF_status_dropdown & "*")
		Call write_bullet_and_variable_in_case_note("Employer Name", evf_employer)
		Call write_bullet_and_variable_in_case_note("EVF for HH member", left(evf_client, 2))
		'for additional information needed
		IF evf_info = "yes" then
			Call write_bullet_and_variable_in_case_note("Additional Info requested: ", evf_info & "- on " & evf_info_date & " by " & evf_request_info)
			If EVF_TIKL_checkbox = checked then call write_variable_in_CASE_NOTE("* TIKL'd for 10 day return.")
		Else
			Call write_variable_in_CASE_NOTE("* No additional information is needed/requested.")
		END IF
		CALL write_bullet_and_variable_in_case_note("Actions taken: ", evf_actions_taken)
		Call write_variable_in_case_note("---")
	End If

	If form_type_array(form_type_const, each_case_note) = iaa_form_name Then 		'IAA Case Notes
		Call start_a_blank_case_note
		If iaa_form_received_checkbox = checked and iaa_ssi_form_received_checkbox = checked Then CALL write_variable_in_case_note("*** IAA and IAA-SSI FORMS RECEIVED ***")
		If iaa_form_received_checkbox = unchecked and iaa_ssi_form_received_checkbox = checked Then CALL write_variable_in_case_note("*** IAA-SSI FORM RECEIVED ***")
		If iaa_form_received_checkbox = checked and iaa_ssi_form_received_checkbox = unchecked Then CALL write_variable_in_case_note("*** IAA FORM RECEIVED ***")
		CALL write_bullet_and_variable_in_case_note("Date Received", iaa_date_received)
		CALL write_bullet_and_variable_in_case_note("Household Member", left(iaa_member_dropdown,2)) 	
		If iaa_form_received_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE("IAA Assistance Type", iaa_type_assistance)
		If iaa_ssi_form_received_checkbox = checked Then CALL write_bullet_and_variable_in_CASE_NOTE("IAA-SSI Interim Assistance", iaa_ssi_type_assistance)
		If iaa_benefits_1 <> "" OR iaa_benefits_2 <> "" OR iaa_benefits_3 <> "" OR iaa_benefits_4 <> "" Then 
			Call write_variable_in_case_note("Other benefits resident may be eligible for: ")
			CALL write_bullet_and_variable_in_case_note("- ", iaa_benefits_1)
			CALL write_bullet_and_variable_in_CASE_NOTE("- ", iaa_benefits_2)
			CALL write_bullet_and_variable_in_CASE_NOTE("- ", iaa_benefits_3)
			CALL write_bullet_and_variable_in_CASE_NOTE("- ", iaa_benefits_4)
		End If 
		If pben_updated = TRUE Then
			CALL write_variable_in_case_note("* PBEN Panel updated")
			CALL write_bullet_and_variable_in_CASE_NOTE(" - Benefit type", iaa_benefit_type)
			CALL write_bullet_and_variable_in_CASE_NOTE(" - Verification", iaa_verification_dropdown)
			CALL write_bullet_and_variable_in_CASE_NOTE(" - Disposition Code", iaa_disposition_code_dropdown)
			CALL write_bullet_and_variable_in_CASE_NOTE(" - Date Applied PBEN", iaa_date_applied_pben)
			CALL write_bullet_and_variable_in_CASE_NOTE(" - Referral Date", iaa_referral_date)
			CALL write_bullet_and_variable_in_CASE_NOTE(" - IAA Date", iaa_iaa_date)
		Else 
			CALL write_variable_in_case_note("* PBEN Panel NOT updated")
		End If
		CALL write_bullet_and_variable_in_case_note("Notes", iaa_comments)
		Call write_variable_in_case_note("---")
	End If
	
	If form_type_array(form_type_const, each_case_note) = mof_form_name Then 		'MOF Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** MEDICAL OPINION FORM RECEIVED ***")
		CALL write_variable_in_CASE_NOTE("* Date Received " & mof_date_received & " for M" & mof_hh_memb)
		IF mof_clt_release_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Client signed release on MOF.")
		Call write_bullet_and_variable_in_case_note("Date of last examination", mof_last_exam_date)
		Call write_bullet_and_variable_in_case_note("Doctor signed form",  mof_doctor_date)
		Call write_bullet_and_variable_in_case_note("Condition will last", mof_time_condition_will_last)
		Call write_bullet_and_variable_in_case_note("Ability to work", mof_ability_to_work)
		Call write_bullet_and_variable_in_case_note("Other notes", mof_other_notes)
		Call write_bullet_and_variable_in_case_note("Actions taken", mof_actions_taken)
		If mof_SSA_application_indicated_checkbox = checked Then Call write_variable_in_CASE_NOTE("* The MOF indicates the client needs to apply for SSA.")
		If mof_TTL_to_update_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Specialized TTL team will review MOF and update the DISA panel as needed.")
		If MOF_TTL_email_checkbox = checked Then Call write_variable_in_CASE_NOTE("* An email regarding this MOF was sent to the TTL/FSS DataTeam for review on " & mof_TTL_email_date & " by " & worker_signature & ".")
		Call write_variable_in_case_note("---")
	End If

	If form_type_array(form_type_const, each_case_note) = mtaf_form_name Then 		'MTAF Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** MINNESOTA TRANSITION APPLICATION RECEIVED ***")
		CALL write_bullet_and_variable_in_CASE_NOTE ("Date received", MTAF_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Date of eligibility", MTAF_MFIP_elig_date)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Address change", mtaf_ADDR_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Household composition change", mtaf_HHcomp_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Change in assets", mtaf_asset_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Change in earned income", mtaf_earned_income_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Change in unearned income", mtaf_unearned_income_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Change in shelter costs", mtaf_shelter_costs_change)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Is housing subsidized? If so, what is the amount", mtaf_subsidized_housing)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Subsidized housing status", mtaf_sub_housing_droplist)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Child or adult care costs", mtaf_child_adult_care_costs)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Proof of relationship on file", mtaf_relationship_proof)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Referred to apply for OMB/PBEN", mtaf_referred_to_OMB_PBEN)
		CALL write_bullet_and_variable_in_CASE_NOTE ("ELIG results fiated", mtaf_elig_results_fiated)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Other notes", mtaf_other_notes)
		CALL write_bullet_and_variable_in_CASE_NOTE ("Verifications Needed", mtaf_verifications_needed)
		If mtaf_signed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MTAF was signed.")
		If mtaf_mfip_financial_orientation_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MFIP orientation information reviewed/completed.")
		If mtaf_ES_exemption_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Client is exempt from cooperation with ES at this time.")
		CALL write_bullet_and_variable_in_CASE_NOTE ("MTAF Status", MTAF_status_dropdown)
		Call write_variable_in_case_note("---")
	End If

	If form_type_array(form_type_const, each_case_note) = psn_form_name Then 		'PSN Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** PROFESSIONAL STATEMENT OF NEED RECEIVED ***")
		CALL write_bullet_and_variable_in_case_note("Date Received", psn_date_received)
		CALL write_bullet_and_variable_in_case_note("Member", left(psn_member_dropdown,2))
		If (psn_section_1_dropdown <> "No- Section NOT completed" OR psn_section_2_dropdown <> "No- Section NOT completed" OR psn_section_3_dropdown <> "No- Section NOT completed" OR psn_section_4_dropdown <> "No- Section NOT completed" OR psn_section_5_dropdown <> "No- Section NOT completed") Then 
			If (psn_section_1_dropdown <> "No- Section NOT completed" OR psn_section_2_dropdown <> "No- Section NOT completed") Then 
				Call write_variable_in_case_note("* The PSN meets GA and GRH basis of eligibility for MB" & (left(psn_member_dropdown, 2)) & " due to their")
			Else 
				Call write_variable_in_case_note("* The PSN meets GRH basis of eligibility for MB" & (left(psn_member_dropdown, 2)) & " due to their") 
			End If
			If (psn_section_1_dropdown = "Yes- At least 1 selected") OR (psn_section_3_dropdown = "Yes- At least 1 selected") OR (psn_section_4_dropdown = "Yes- At least 2 selected") Then CALL write_variable_in_case_note("    -needed assistance to access or maintain housing")
			If psn_section_2_dropdown = "Yes- 1 selected" Then CALL write_variable_in_case_note( "    -disabling condition")
			If psn_section_5_dropdown = "Yes- Section completed" Then CALL write_variable_in_case_note("    -exit of a residential behavioral health treatment with instable housing")
		End If
		CALL write_variable_in_case_note("* PSN Signed by " & psn_cert_prof & " at " & psn_facility)
		CALL write_bullet_and_variable_in_case_note("Section 1: Housing Situation", psn_section_1_dropdown)
		CALL write_bullet_and_variable_in_case_note("Section 2: Disabling Condtion", psn_section_2_dropdown)
		CALL write_bullet_and_variable_in_case_note("Section 3: MA Housing Stabilization Services", psn_section_3_dropdown)
		CALL write_bullet_and_variable_in_case_note("Section 4: MN Housing Support Supplemental Services", psn_section_4_dropdown)
		CALL write_bullet_and_variable_in_case_note("Section 5: Transition from Residential Treatment to MN HS Program", psn_section_5_dropdown)
		If psn_udpate_wreg_disa_checkbox = checked Then 
			CALL write_variable_in_case_note("* WREG and DISA panels updated")
		Else
			CALL write_variable_in_case_note("* WREG and DISA panels NOT updated")
		End If
		CALL write_bullet_and_variable_in_case_note("Comments", psn_comments)
		Call write_variable_in_case_note("---")
	End If
	
	If form_type_array(form_type_const, each_case_note) = sf_form_name Then 		'SF Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** SHELTER FORM RECEIVED ***")
		CALL write_bullet_and_variable_in_case_note("Form Name",sf_name_of_form)
		CALL write_bullet_and_variable_in_case_note("Date Received", sf_date_received)
		CALL write_bullet_and_variable_in_case_note("Tenant Name", sf_tenant_name)
		CALL write_bullet_and_variable_in_case_note("Total Rent", sf_total_rent)
		If sf_adults <> "" or sf_children <> "" Then CALL write_variable_in_case_note("* Person(s) in Unit")
		CALL write_bullet_and_variable_in_case_note("  Adults", sf_adults)
		CALL write_bullet_and_variable_in_CASE_NOTE ("  Children", sf_children)
		CALL write_bullet_and_variable_in_case_note("  Street Address", sf_street_address)
		CALL write_bullet_and_variable_in_case_note("  City", sf_city)
		CALL write_bullet_and_variable_in_case_note("  State", sf_state)
		CALL write_bullet_and_variable_in_case_note("  Zip", sf_zip)
		If (sf_electrcity_checkbox = checked OR sf_ac_checkbox = checked OR sf_heat_checkbox = checked OR sf_gas_checkbox = checked OR sf_garbage_removal_checkbox = checked OR sf_water_sewer_checkbox = checked OR sf_garage_plugin_checkbox = checked OR sf_none_checkbox = checked OR sf_other_checkbox = checked) Then CALL write_variable_in_case_note("* Utilities tenant is responsible to pay separate from the rent amount:")
		If sf_electrcity_checkbox = checked Then CALL write_variable_in_case_note("  - Electrcity") 
		If sf_ac_checkbox = checked Then CALL write_variable_in_case_note("  - Air Conditioning") 
		If sf_heat_checkbox = checked Then CALL write_variable_in_case_note("  - Heat") 
		If sf_gas_checkbox = checked Then CALL write_variable_in_case_note("  - Gas") 
		If sf_garbage_removal_checkbox = checked Then CALL write_variable_in_case_note("  - Garbage removal") 
		If sf_water_sewer_checkbox = checked Then CALL write_variable_in_case_note("  - Water and sewer") 
		If sf_garage_plugin_checkbox = checked Then CALL write_variable_in_case_note("  - Garage/plug-in") 
		If sf_none_checkbox = checked Then CALL write_variable_in_case_note("  - None") 
		If sf_other_checkbox = checked Then CALL write_variable_in_case_note("  - Other: " & sf_other_utility) 
		If sf_update_addr_shel_hest_checkbox Then 
			If addr_update_attempted = true Then CALL write_variable_in_case_note("* ADDR panel updated")
			If hest_updated = True Then CALL write_variable_in_case_note("* HEST panel updated")
			If shel_updated = True Then CALL write_variable_in_case_note("* SHEL panel updated")
		End If
		CALL write_bullet_and_variable_in_case_note("Comments", sf_comments)
		If sf_tikl_nav_check = checked Then write_variable_in_case_note("* TIKL set")
		Call write_variable_in_case_note("---")		
	End If

	If form_type_array(form_type_const, each_case_note) = diet_form_name Then 		'Special Diet Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** SPECIAL DIET FORM RECEIVED ***")	
		CALL write_bullet_and_variable_in_case_note("Date Received", diet_date_received)	
		CALL write_bullet_and_variable_in_case_note("Member", diet_ref_number)							'required
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

		If diet_mfip_msa_status = "Not Active/Pending" Then CALL write_variable_in_case_note("* DIET panel NOT updated- case is not active/pending for MSA or MFIP")
		If diet_status_dropdown = "Incomplete" then
			CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown & "- form returned to client")
		ElseIf  diet_status_dropdown = "Denied" Then
			CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown & "- The doctor has NOT indicated an eligible diet need.")
		Else
			CALL write_bullet_and_variable_in_case_note("Diet status", diet_status_dropdown)
			CALL write_bullet_and_variable_in_CASE_NOTE("Updated DIET panel for months", updated_diet_months)

		End If 
		CALL write_bullet_and_variable_in_case_note("Comments",diet_comments)
		CALL write_variable_in_case_note("---")
	End If

	If form_type_array(form_type_const, each_case_note) = other_form_name Then 		'Other Case Notes
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** Other forms received: " & other_list_form_names)	
		CALL write_bullet_and_variable_in_case_note("Date Received", other_date_received)
		CALL write_bullet_and_variable_in_case_note("Document Notes", other_doc_notes)
		CALL write_bullet_and_variable_in_case_note("Verifications Received", other_verif_received)
		CALL write_bullet_and_variable_in_case_note("Action Taken", other_action_taken)
		CALL write_variable_in_case_note("---")
	End If 
	If each_case_note = Ubound(form_type_array, 2) Then CALL write_variable_in_case_note(worker_signature)
Next

'For/Next creates individual case notes for the following documents. Casenoting these individually so we can search for them in the future.
For unique_case_notes = 0 to Ubound(form_type_array, 2)			
	If form_type_array(form_type_const, unique_case_notes) = atr_form_name Then 'ATR Case Notes
		PF3
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** ATR RECEIVED *** FOR M" & left(atr_member_dropdown, 2) & " - " & atr_name & " - Release Ends: " & atr_end_date)
		CALL write_bullet_and_variable_in_case_note("Date Received", atr_date_received)
		CALL write_bullet_and_variable_in_case_note("Member", left(atr_member_dropdown, 2))
		CALL write_bullet_and_variable_in_case_note("Start Date", atr_start_date)
		CALL write_bullet_and_variable_in_case_note("End Date", atr_end_date)
		CALL write_bullet_and_variable_in_case_note("Authorization Type", atr_authorization_type)
		CALL write_bullet_and_variable_in_case_note("Contact Type", atr_contact_type)
		CALL write_bullet_and_variable_in_case_note("  Contact Name", atr_name)
		CALL write_bullet_and_variable_in_case_note("  Phone Number", atr_phone_number)
		CALL write_bullet_and_variable_in_case_note("  Fax Number", atr_fax_number)
		CALL write_bullet_and_variable_in_case_note("  Email", atr_email)

		If atr_eval_treat_checkbox = checked Then CALL write_variable_in_case_note("* Record requested will be used to continue evaluation or treatment")
		If atr_coor_serv_checkbox = checked Then CALL write_variable_in_case_note("* Record requested will be used to coordinate services")
		If atr_elig_serv_checkbox = checked Then CALL write_variable_in_case_note("* Record requested will be used to determine eligibility for assistance/service")
		If atr_court_checkbox = checked Then CALL write_variable_in_case_note("* Record requested will be used for court proceedings")
		If atr_other_checkbox = checked Then CALL write_bullet_and_variable_in_case_note("Record requested will be used", atr_other)
		CALL write_bullet_and_variable_in_case_note("Comments", atr_comments)
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)
	End If

	If form_type_array(form_type_const, unique_case_notes) = arep_form_name then 'AREP Case Notes
		PF3
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** AREP Received ***")
		call write_variable_in_CASE_NOTE("* Received: " & AREP_recvd_date & ". AREP: " & arep_name)
		If arep_dhs_3437_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the DHS 3437 - MHCP AUTHORIZED REPRESENTATIVE REQUEST Form.")
		If arep_HC_12729_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the HC 12729 - AUTHORIZED REPRESENTATIVE REQUEST Form.")
		If arep_D405_checkbox = checked Then
			Call write_variable_in_CASE_NOTE("  - AREP name on the SNAP AUTHORIZED REPRESENTATIVE CHOICE D405 Form.")
			Call write_variable_in_CASE_NOTE("  - AREP also authorized to get and use EBT Card.")
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
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)
	End If
	
	If form_type_array(form_type_const, unique_case_notes) = hosp_form_name Then 'Hospice Case Notes
		PF3
		Call start_a_blank_case_note
		Call write_variable_in_case_note("*** HOSPICE TRANSACTION FORM RECEIVED ***") 'DO NOT cchange name for Hospice.Must keep the same header otherwise reading of past case notes won't work/continue 
		Call write_bullet_and_variable_in_CASE_NOTE("Client", hosp_resident_name)
		Call write_bullet_and_variable_in_CASE_NOTE("Hospice Name", hosp_name)
		Call write_bullet_and_variable_in_CASE_NOTE("NPI Number", hosp_npi_number)
		Call write_bullet_and_variable_in_CASE_NOTE("Date of Entry", hosp_entry_date)
		Call write_bullet_and_variable_in_CASE_NOTE("Exit Date", hosp_exit_date)
		Call write_bullet_and_variable_in_CASE_NOTE("MMIS updated as of", hosp_mmis_updated_date)
		Call write_bullet_and_variable_in_CASE_NOTE("MMIS not updated due to", hosp_reason_not_updated)
		Call write_bullet_and_variable_in_CASE_NOTE("Notes", hosp_other_notes)
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)
	End If

	If form_type_array(form_type_const, unique_case_notes) = ltc_1503_form_name Then 'LTC 1503 Case Notes
		PF3
		Call start_a_blank_case_note
		CALL write_variable_in_case_note("*** LTC-1503 FORM RECEIVED ***")
		If ltc_1503_processed_1503_checkbox = checked then
			call write_variable_in_CASE_NOTE("***Processed 1503 from " & ltc_1503_FACI_1503 & "***")
		Else
			call write_variable_in_CASE_NOTE("***Rec'd 1503 from " & ltc_1503_FACI_1503 & ", DID NOT PROCESS***")
		End if
		If ltc_1503_FACI_update_checkbox = checked Then Call write_variable_in_case_note("* Updated FACI for " & ltc_1503_faci_footer_month & "/" & ltc_1503_faci_footer_year)
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
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)
	End If
Next

'Documenting verifications still outstanding 
If ButtonPressed = none_btn OR trim(outstanding_verifs) <> "" Then 
	PF3
	Call start_a_blank_case_note
	CALL write_bullet_and_variable_in_case_note("Outstanding Verifications", outstanding_verifs)
	CALL write_variable_in_case_note(worker_signature)
End If

script_end_procedure_with_error_report("Success! " & vbcr & end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/20/2024
'--Tab orders reviewed & confirmed----------------------------------------------06/20/2024
'--Mandatory fields all present & Reviewed--------------------------------------06/20/2024
'--All variables in dialog match mandatory fields-------------------------------06/20/2024
'Review dialog names for content and content fit in dialog----------------------06/20/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------06/20/2024
'--Include script category and name somewhere on first dialog-------------------06/20/2024
'--Create a button to reference instructions------------------------------------06/20/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/20/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------06/20/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/20/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------06/20/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/20/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------06/20/2024
'--PRIV Case handling reviewed -------------------------------------------------06/20/2024
'--Out-of-County handling reviewed----------------------------------------------NA
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/20/2024
'--BULK - review output of statistics and run time/count (if applicable)--------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------06/20/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------NA
'--BULK - remove 1 incrementor at end of script reviewed------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/20/2024
'--comment Code-----------------------------------------------------------------06/20/2024
'--Update Changelog for release/update------------------------------------------06/20/2024
'--Remove testing message boxes-------------------------------------------------06/20/2024
'--Remove testing code/unnecessary code-----------------------------------------06/20/2024
'--Review/update SharePoint instructions----------------------------------------06/20/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/20/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------NA
'--Complete misc. documentation (if applicable)---------------------------------NA
'--Update project team/issue contact (if applicable)----------------------------06/20/2024