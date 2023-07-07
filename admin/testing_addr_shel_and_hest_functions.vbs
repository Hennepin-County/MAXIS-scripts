'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Testing ADDR SHEL and HEST Functions.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 125          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================
run_locally = True
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

'DECLARATIONS ============================================================================================================='
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

Dim ALL_SHEL_PANELS_ARRAY()
ReDim ALL_SHEL_PANELS_ARRAY(shel_entered_notes_const, 0)

ADDR_dlg_page = 1
SHEL_dlg_page = 2
HEST_dlg_page = 3
CHNG_page_btn = 4

ADDR_page_btn = 100
SHEL_page_btn = 101
HEST_page_btn = 102
CHNG_dlg_page = 103

update_information_btn 	= 500
save_information_btn	= 501
clear_mail_addr_btn		= 502
clear_phone_one_btn		= 503
clear_phone_two_btn		= 504
clear_phone_three_btn	= 505
clear_all_btn			= 506
view_total_shel_btn		= 507
update_household_percent_button = 508
housing_change_continue_btn = 509
housing_change_overview_btn = 510
housing_change_addr_update_btn = 511
housing_change_shel_update_btn = 512
housing_change_shel_details_btn = 513

enter_shel_one_btn = 550
enter_shel_two_btn = 551
enter_shel_three_btn = 552

update_addr = False
update_shel = False
update_hest = False
caf_answer_droplist = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
show_totals = True

total_current_rent 		= 0
total_current_taxes 	= 0
total_current_lot_rent 	= 0
total_current_room 		= 0
total_current_mortgage 	= 0
total_current_garage 	= 0
total_current_insurance = 0
total_current_subsidy 	= 0
total_paid_to = ""
total_paid_by_household = 100
total_paid_by_others = 0

'==========================================================================================================================

EMConnect ""
Call check_for_MAXIS(TRUE)
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
EmWriteScreen MAXIS_case_number, 18, 43

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region <> "TRAINING" Then script_end_procedure("This is a testing script only, for BZST script writers. It should only be run in TRAINING region. The script will now end as this does not appear to be TRAINING region.")

BeginDialog Dialog1, 0, 0, 231, 100, "Dialog"
  EditBox 60, 10, 50, 15, MAXIS_case_number
  EditBox 180, 10, 20, 15, MAXIS_footer_month
  EditBox 205, 10, 20, 15, MAXIS_footer_year
  DropListBox 60, 35, 165, 45, "Application/Renewal"+chr(9)+"Change", select_option
  ButtonGroup ButtonPressed
    OkButton 175, 80, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 130, 15, 50, 10, "Footer Month:"
  Text 10, 40, 50, 10, "Action to Run:"
  Text 5, 55, 220, 20, "NOTE: This script does not have err_msg handling as it is for script writers only. Ensure you have entered ALL above information."
EndDialog

dialog Dialog1


'SEARCH THE LIST OF HOUSEHOLD MEMBERS TO SEARCH ALL SHEL PANELS
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
	'MsgBox access_denied_check
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
	Call navigate_to_MAXIS_screen("STAT", "SHEL")
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

	Call navigate_to_MAXIS_screen("STAT", "MEMB")
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

If select_option = "Application/Renewal" Then
	' Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit,                   mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_addr_panel_info, addr_update_attempted)
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
	' Call reformat_phone_number(phone_two, "(111) 222-3333")

	Call access_HEST_panel("READ", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
	For shel_member = 0 to UBound(ALL_SHEL_PANELS_ARRAY, 2)
		If ALL_SHEL_PANELS_ARRAY(shel_exists_const, shel_member) = True Then
			Call access_SHEL_panel("READ", ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, shel_member), ALL_SHEL_PANELS_ARRAY(hud_sub_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(shared_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(paid_to_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member))


			' total_current_rent 		= total_current_rent + 		ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, 	shel_member)
			' total_current_taxes 	= total_current_taxes + 	ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, 		shel_member)
			' total_current_lot_rent 	= total_current_lot_rent + 	ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member)
			' total_current_room 		= total_current_room + 		ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, 	shel_member)
			' total_current_mortgage 	= total_current_mortgage + 	ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member)
			' total_current_garage 	= total_current_garage + 	ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, 	shel_member)
			' total_current_insurance = total_current_insurance + ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const,shel_member)
			' total_current_subsidy 	= total_current_subsidy + 	ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, 	shel_member)
		End If
	Next
	page_to_display = ADDR_dlg_page
	Call read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, total_shelter_expense, total_shel_original_information)

	addr_update_attempted = False
	shel_update_attempted = False
	hest_update_attempted = False

	 resi_line_one = ""
	 resi_line_two = ""
	 mail_line_one = ""
	 mail_line_two = ""

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
		cancel_without_confirmation

		If page_to_display = ADDR_dlg_page Then Call navigate_ADDR_buttons(update_addr, err_msg, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, mail_street_full, mail_city, mail_state, mail_zip, phone_one, phone_two, phone_three, type_one, type_two, type_three)
		If page_to_display = SHEL_dlg_page Then Call navigate_SHEL_buttons(update_shel, show_totals, err_var, ALL_SHEL_PANELS_ARRAY, member_selection, shel_ref_number_const, shel_exists_const, hud_sub_yn_const, shared_yn_const, paid_to_const, rent_retro_amt_const, rent_retro_verif_const, rent_prosp_amt_const, rent_prosp_verif_const, lot_rent_retro_amt_const, lot_rent_retro_verif_const, lot_rent_prosp_amt_const, lot_rent_prosp_verif_const, mortgage_retro_amt_const, mortgage_retro_verif_const, mortgage_prosp_amt_const, mortgage_prosp_verif_const, insurance_retro_amt_const, insurance_retro_verif_const, insurance_prosp_amt_const, insurance_prosp_verif_const, tax_retro_amt_const, tax_retro_verif_const, tax_prosp_amt_const, tax_prosp_verif_const, room_retro_amt_const, room_retro_verif_const, room_prosp_amt_const, room_prosp_verif_const, garage_retro_amt_const, garage_retro_verif_const, garage_prosp_amt_const, garage_prosp_verif_const, subsidy_retro_amt_const, subsidy_retro_verif_const, subsidy_prosp_amt_const, subsidy_prosp_verif_const, update_information_btn, save_information_btn, memb_btn_const, attempted_update_const, clear_all_btn, view_total_shel_btn, update_household_percent_button)

		If page_to_display = HEST_dlg_page Then Call navigate_HEST_buttons(update_hest, err_msg, update_information_btn, save_information_btn, choice_date, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, date)
		If err_msg <> "" then MsgBox "Please Resolve:" & vbCr & err_msg

		If ButtonPressed = ADDR_page_btn Then page_to_display = ADDR_dlg_page
		If ButtonPressed = SHEL_page_btn Then page_to_display = SHEL_dlg_page
		If ButtonPressed = HEST_page_btn Then page_to_display = HEST_dlg_page
	Loop until ButtonPressed = -1

	' If addr_update_attempted = True Then
	addr_eff_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year
	' Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit,                   mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_addr_panel_info, addr_update_attempted)
	Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
	For shel_member = 0 to UBound(ALL_SHEL_PANELS_ARRAY, 2)
		' If ALL_SHEL_PANELS_ARRAY(attempted_update_const, shel_member) = True Then
		Call access_SHEL_panel("WRITE", ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, shel_member), ALL_SHEL_PANELS_ARRAY(hud_sub_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(shared_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(paid_to_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member))
		' End If
	Next
	' If hest_update_attempted = True Then
	Call access_HEST_panel("WRITE", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)


	end_msg = "Finished the 'Application/Renewal' option which tests the panel displays, navigation, and update of ADRR, SHEL, and HEST. The panels should be updated. To view the question workflow created for housing change options, select 'Change' at the beginning of this script run."
End If


If select_option = "Change" Then

	Call read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, total_shelter_expense, total_shel_original_information)
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
	Call access_HEST_panel("READ", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

	original_resi_street_full = resi_street_full
	original_resi_city = resi_city
	original_resi_state = resi_state
	original_resi_zip = resi_zip

	housing_questions_step = 1
	view_addr_update_dlg = False
	view_shel_update_dlg = False
	view_shel_details_dlg = False
	page_to_display = CHNG_dlg_page

	Do
		err_msg = ""

		' MsgBox "HOUSING QUESTIONS - " & housing_questions_step & vbCr & "SHEL UPDATE - " & shel_update_step

		BeginDialog Dialog1, 0, 0, 555, 385, "Housing Expense Detail"

		  ButtonGroup ButtonPressed

			If page_to_display = CHNG_dlg_page Then
				Text 503, 57, 60, 10, "CHANGE"
				Call display_HOUSING_CHANGE_information(housing_questions_step, shel_update_step, notes_on_address, original_resi_street_full, original_resi_city, original_resi_state, original_resi_zip, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, ALL_SHEL_PANELS_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, hh_comp_change, hest_heat_ac_yn_code, hest_electric_yn_code, hest_phone_yn_code, new_total_shelter_expense_amount, people_paying_SHEL, verif_detail, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)

			End If

			If page_to_display <> CHNG_dlg_page Then PushButton 485, 55, 65, 13, "CHANGE", CHNG_page_btn

			OkButton 450, 365, 50, 15
			CancelButton 500, 365, 50, 15

		EndDialog


		Dialog Dialog1
		cancel_without_confirmation
		' MsgBox "Button - " & ButtonPressed

		If ButtonPressed = -1 AND housing_questions_step <> 5 Then ButtonPressed = housing_change_continue_btn
		If page_to_display = CHNG_dlg_page Then Call navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, ALL_SHEL_PANELS_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_total_shelter_expense_amount, people_paying_SHEL, verif_detail, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)

		If ButtonPressed = CHNG_page_btn Then page_to_display = CHNG_dlg_page
		If err_msg <> "" then MsgBox "Please Resolve:" & vbCr & err_msg
	Loop until ButtonPressed = -1
	end_msg = "Finished the Housing Change function display. This function does not update panels as it is intended to show the dialog display only. The update of panels would be added using the 'access_PANEL' functions after this dialog is displayed. To test the updating capabilities, use the 'Application/Renewal' option at the beginning of this script."
End If


script_end_procedure(end_msg)
