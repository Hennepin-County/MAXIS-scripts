'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - HOUSING DETAIL UPDATE.vbs"
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


'FUNCTIONS ================================================================================================================
function display_HOUSING_CHANGE_information(housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, THE_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)


	yes_no_list = "?"+chr(9)+"Yes"+chr(9)+"No"
	x_pos = 345
	If view_shel_details_dlg = True Then
		shel_det_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	If view_shel_update_dlg = True Then
		shel_up_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	If view_addr_update_dlg = True Then
		addr_up_x_pos = x_pos
		x_pos = x_pos - 60
	End If
	overview_x_pos = x_pos

	GroupBox 10, 10, 460, 355, "Change in HOUSING Information"

	If housing_questions_step = 1 Then
		Text overview_x_pos + 10, 10, 60, 13, "OVERVIEW"

		GroupBox 15, 25, 450, 75, "Address"
		Text 75, 45, 95, 10, "Did the household move?"
		DropListBox 170, 40, 45, 45, yes_no_list, household_move_yn
		Text 25, 65, 145, 10, "Did everyone in the household move with?"
		DropListBox 170, 60, 45, 45, yes_no_list, household_move_everyone_yn
		Text 75, 85, 90, 10, "What date did they move?"
		EditBox 170, 80, 45, 15, move_date
		Text 255, 40, 95, 10, "Current Residence Address:"
		Text 265, 55, 190, 10, resi_street_full
		Text 265, 70, 190, 10, resi_city & ", " & left(resi_state, 2) & " " & resi_zip
		' PushButton 390, 85, 70, 10, "UPDATE ADDR", update_addr_button

		GroupBox 15, 105, 450, 70, "Housing Cost"
		Text 40, 125, 130, 10, "Is there a change to the housing cost?"
		DropListBox 170, 120, 45, 45, yes_no_list, shel_change_yn

		Text 40, 145, 125, 10, "What date did the new expense start?"
		EditBox 170, 140, 45, 15, shel_start_date

		Text 265, 115, 95, 10, "Current Housing Costs"
		Text 280, 130, 35, 10, " Rent: "
		Text 305, 130, 30, 10, "$ " & total_current_rent
		Text 375, 130, 40, 10, " Taxes: "
		Text 405, 130, 30, 10, "$ " & total_current_taxes
		Text 270, 140, 45, 10, "Lot Rent: "
		Text 305, 140, 30, 10, "$ " & total_current_lot_rent
		Text 375, 140, 40, 10, " Room: "
		Text 405, 140, 30, 10, "$ " & total_current_room
		Text 265, 150, 50, 10, " Mortgage: "
		Text 305, 150, 30, 10, "$ " & total_current_mortgage
		Text 370, 150, 45, 10, " Garage: "
		Text 405, 150, 30, 10, "$ " & total_current_garage
		Text 265, 160, 50, 10, "Insurance: "
		Text 305, 160, 30, 10, "$ " & total_current_insurance
		Text 370, 160, 45, 10, "Subsidy: "
		Text 405, 160, 30, 10, "$ " & total_current_subsidy
		' Text 270, 185, 95, 10, "New shelter expense is a(n)"
		' DropListBox 365, 180, 95, 45, ""+chr(9)+"Increate"+chr(9)+"Decrease"+chr(9)+"No Difference", shel_change_type
		' PushButton 390, 205, 70, 10, "UPDATE SHEL", update_shel_button

		GroupBox 15, 180, 450, 115, "Utilities Expense"
	    Text 25, 195, 275, 10, "Is the household responsible to paythe Heat Expense or Air Conditioner Expense?"
	    DropListBox 295, 190, 45, 45, yes_no_list, hest_heat_ac_yn
	    Text 25, 215, 180, 10, "Is the household responsible to pay electric expense?"
	    DropListBox 210, 210, 45, 45, yes_no_list, hest_electric_yn
	    Text 40, 230, 235, 10, "If yes, is there any AC plugged into that is used at any point in the year?"
	    DropListBox 280, 225, 45, 45, yes_no_list, hest_ac_on_electric_yn
	    Text 40, 250, 235, 10, "If yes, does this include any heat source during any point in the year?"
	    DropListBox 280, 245, 45, 45, yes_no_list, hest_heat_on_electric_yn
	    Text 25, 270, 145, 10, "Is anyone responsible to PAY for a phone?"
	    DropListBox 170, 265, 45, 45, yes_no_list, hest_phone_yn
	    Text 30, 280, 230, 10, "(Free phone plans without a payment requirement cannot be counted.)"

	End If

	' view_addr_update_dlg
	' view_shel_update_dlg
	' view_shel_details_dlg

	If housing_questions_step = 2 Then
		Text addr_up_x_pos + 5, 10, 60, 10, "ADDR UPDATE"

		Text 15, 25, 450, 10, "STEP 2 - ADDR UPDATES  -  Enter new address information here:"

		Call display_ADDR_information(True, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)

		' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	End If
	If housing_questions_step = 3 Then
		Text shel_up_x_pos + 5, 10, 60, 10, "SHEL UPDATE"

		Text 15, 25, 450, 10, "STEP 3 - SHEL UPDATES"

		' what_is_the_living_arrangement
		' unit_owned
		' new_total_rent_amount
		' new_rent_subsidy_yn
		' new_renter_insurance_amount
		' new_renters_insurance_required_yn
		' new_total_garage_rent_amount
		' new_garage_rent_required_yn
		' new_rent_paid_to_name
		' other_person_checkbox
		' other_person_name
		' payment_split_evenly_yn
		' ALL_SHEL_PANELS_ARRAY
		' person_age_const
		' person_shel_checkbox
		' shel_ref_number_const
		' new_shel_pers_total_amt_const
		' new_shel_pers_total_amt_type_const
		' other_new_shel_total_amt
		' other_new_shel_total_amt_type

		If shel_update_step > 0 Then
			Text 20, 45, 95, 10, "What is the living situation?"
		    DropListBox 115, 40, 125, 45, "Select One..."+chr(9)+"Apartment or Townhouse"+chr(9)+"House"+chr(9)+"Trailer Home/Mobile Home"+chr(9)+"Room Only"+chr(9)+"Shelter"+chr(9)+"Hotel"+chr(9)+"Vehicle"+chr(9)+"Other", what_is_the_living_arrangement
		    Text 250, 45, 120, 10, "Does the household own the home?"
		    DropListBox 370, 40, 90, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", unit_owned
		    PushButton 410, 55, 50, 10, "Enter", enter_shel_one_btn
		End If

		If shel_update_step > 1 Then
			GroupBox 15, 65, 450, 155, "Payment Details"
			If (what_is_the_living_arrangement = "Apartment or Townhouse" OR what_is_the_living_arrangement = "House") Then
				If unit_owned = "No" Then
				    Text 20, 80, 105, 10, "What is the total rent amount?"
				    EditBox 130, 75, 50, 15, new_total_rent_amount
					Text 225, 80, 100, 10, "Who is the expense paid to?"
					EditBox 325, 75, 135, 15, new_SHEL_paid_to_name
				    Text 20, 100, 195, 10, "Does the household receive a subsidy for the rent amount?"
				    DropListBox 220, 95, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_rent_subsidy_yn
					Text 290, 100, 75, 10, "Subsidy Amount:"
					EditBox 365, 95, 50, 15, new_total_subsidy_amount
				    Text 20, 120, 150, 10, "What is the amount of any renters insurance?"
				    EditBox 175, 115, 50, 15, new_renter_insurance_amount
				    Text 260, 120, 135, 10, "Is this insurance required per the lease?"
				    DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_renters_insurance_required_yn
				    Text 20, 140, 130, 10, "What is the amount of the garage rent?"
				    EditBox 150, 135, 50, 15, new_total_garage_amount
				    Text 250, 140, 145, 10, "Is this garage rental required per the lease?"
				    DropListBox 400, 135, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_garage_rent_required_yn

				End If

				If unit_owned = "Yes" Then
					Text 20, 80, 115, 10, "What is the total mortgage amount?"
					EditBox 140, 75, 50, 15, new_total_mortgage_amount
					Text 230, 80, 170, 10, "Does this payment include all taxes and insturance?"
					DropListBox 400, 75, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_mortgage_have_escrow_yn
					Text 20, 100, 150, 10, "How much insurance do you pay seperately?"
					EditBox 170, 95, 40, 15, new_morgage_insurance_amount
					Text 220, 100, 195, 10, "Do have more insurance than required by the mortgage?"
					DropListBox 400, 95, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_excess_insurance_yn
					Text 20, 120, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 115, 50, 15, new_total_tax_amount
					' Text 250, 120, 145, 10, "Is this garage rental required per the lease?"
					' DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_garage_rent_required_yn
					Text 20, 140, 100, 10, "Who is the mortgage paid to?"
					EditBox 120, 135, 135, 15, new_SHEL_paid_to_name
				End If
			ElseIf what_is_the_living_arrangement = "Trailer Home/Mobile Home" Then
				If unit_owned = "No" Then
					Text 20, 80, 110, 10, "What is the total unit rent amount?"
					EditBox 135, 75, 50, 15, new_total_rent_amount
					Text 20, 100, 105, 10, "What is the lot rent Amount?"
					EditBox 130, 95, 50, 15, new_total_lot_rent_amount
					Text 20, 120, 150, 10, "What is the amount of any renters insurance?"
					EditBox 175, 115, 50, 15, new_renter_insurance_amount
					Text 260, 120, 135, 10, "Is this insurance required per the lease?"
					DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_renters_insurance_required_yn

					Text 20, 140, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 135, 50, 15, new_total_tax_amount
					' Text 250, 120, 145, 10, ""
					' DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_garage_rent_required_yn
					Text 225, 140, 100, 10, "Who is the expense paid to?"
					EditBox 325, 135, 135, 15, new_SHEL_paid_to_name
				End If

				If unit_owned = "Yes" Then
					Text 20, 80, 115, 10, "What is the total mortgage amount?"
					EditBox 140, 75, 50, 15, new_total_mortgage_amount
					Text 230, 80, 170, 10, "Does this payment include all taxes and insturance?"
					DropListBox 400, 75, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_mortgage_have_escrow_yn
					Text 20, 100, 105, 10, "What is the lot rent Amount?"
					EditBox 135, 95, 50, 15, new_total_lot_rent_amount
					Text 20, 120, 150, 10, "How much insurance do you pay seperately?"
					EditBox 170, 115, 40, 15, new_morgage_insurance_amount
					Text 220, 120, 195, 10, "Do have more insurance than required by the mortgage?"
					DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_excess_insurance_yn
					Text 20, 140, 140, 10, "How much in taxes do you pay seperately?"
					EditBox 160, 135, 50, 15, new_total_tax_amount
					' Text 250, 120, 145, 10, ""
					' DropListBox 400, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_garage_rent_required_yn
					Text 225, 140, 100, 10, "Who is the expense paid to?"
					EditBox 325, 135, 135, 15, new_SHEL_paid_to_name
				End If
			ElseIf what_is_the_living_arrangement = "Room Only" Then
				Text 20, 80, 115, 10, "What is the total room rent amount?"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 100, 10, "Who is the room rent paid to?"
				EditBox 120, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Shelter" Then
				Text 20, 80, 115, 10, "What is the cost for the shelter"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 120, 10, "Who is the shelter expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Hotel" Then
				Text 20, 80, 115, 10, "What is the cost for the hotel room?"
				EditBox 140, 75, 50, 15, new_total_room_amount
				Text 200, 80, 65, 10, "How often paid?"
				DropListBox 265, 75, 60, 45, "Select One..."+chr(9)+"Nightly"+chr(9)+"Weekly"+chr(9)+"Monthly", new_room_payment_frequency
				Text 20, 100, 120, 10, "Who is the hotel expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Vehicle" Then
				Text 20, 80, 115, 10, "How much insurance do you pay?"
				EditBox 135, 75, 40, 15, new_vehicle_insurance_amount
				Text 20, 100, 120, 10, "Who is the vehicle expense paid to?"
				EditBox 140, 95, 135, 15, new_SHEL_paid_to_name
			ElseIf what_is_the_living_arrangement = "Other" Then
				Text 25, 80, 35, 10, "Rent: $"
				EditBox 55, 75, 30, 15, new_total_rent_amount
				Text 120, 80, 60, 10, "Mortgage: $"
				EditBox 165, 75, 30, 15, new_total_mortgage_amount
				Text 230, 80, 50, 10, "Lot Rent: $"
				EditBox 275, 75, 30, 15, new_total_lot_rent_amount
				Text 330, 80, 35, 10, "Room: $"
				EditBox 365, 75, 30, 15, new_total_room_amount

				Text 20, 100, 40, 10, "Taxes: $"
				EditBox 55, 95, 30, 15, new_total_tax_amount
				Text 125, 100, 45, 10, "Garage: $"
				EditBox 165, 95, 30, 15, new_total_garage_amount
				Text 225, 100, 55, 10, "Insurance: $"
				EditBox 275, 95, 30, 15, new_total_insurance_amount
				Text 330, 100, 35, 10, "Subsidy: $"
				EditBox 365, 95, 30, 15, new_total_subsidy_amount
				' Text 20, 120, 195, 10, "Does the household receive a subsidy for the rent amount?"
				' DropListBox 215, 115, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", new_rent_subsidy_yn
				Text 20, 120, 120, 10, "Who is the housing expense paid to?"
				EditBox 140, 115, 135, 15, new_SHEL_paid_to_name
				' Text 90, 160, 45, 10, "Subsidy: $"
				' Text 405, 160, 30, 10, "$ " & total_current_subsidy
			End If
			GroupBox 20, 155, 440, 65, "Check the box for each person responsible for the housing payment:"
			x_pos = 30
			y_pos = 170
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_age_const, the_membs) >= 18 Then
					CheckBox 30, 170, 80, 10, "MEMB " & THE_ARRAY(shel_ref_number_const, the_membs), THE_ARRAY(person_shel_checkbox, the_membs)
					x_pos = x_pos + 125
					If x_pos = 200 Then
						y_pos = y_pos + 15
					End If
				End If
			next
			' CheckBox 30, 170, 80, 10, "Check1", Check1
			' CheckBox 30, 185, 80, 10, "Check2", Check2
			' CheckBox 155, 170, 80, 10, "Check3", Check3
			' CheckBox 155, 185, 80, 10, "Check4", Check4
			CheckBox 290, 170, 145, 10, "Someone outside the household. Name:", other_person_checkbox
			EditBox 305, 180, 150, 15, other_person_name
			Text 25, 205, 200, 10, "Is the payment split evenly among all the responsible parties?"
			DropListBox 230, 200, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", payment_split_evenly_yn
			PushButton 405, 205, 50, 10, "Enter", enter_shel_two_btn
		End If

		If shel_update_step > 2 Then
		    GroupBox 15, 225, 450, 50, "How is the payment split?"
			x_pos = 25
			y_pos = 240
			If new_rent_subsidy_yn = "Yes" Then
				Text x_pos, y_pos, 60, 10, "Subsidy pays: "
				EditBox x_pos + 65, y_pos - 5, 50, 15, new_total_subsidy_amount
				Text x_pos + 120, y_pos, 55, 10, "dollars"
				' DropListBox x_pos + 120, y_pos - 5, 55, 45, "dollars"+chr(9)+"percent", THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs)
				x_pos = x_pos + 195
				If x_pos = 415 Then
					x_pos = 25
					y_pos = y_pos + 20
				End If
			End If
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
					Text x_pos, y_pos, 60, 10, "MEMB " & THE_ARRAY(shel_ref_number_const, the_membs) & " pays: "
					EditBox x_pos + 65, y_pos - 5, 50, 15, THE_ARRAY(new_shel_pers_total_amt_const, the_membs)
					DropListBox x_pos + 120, y_pos - 5, 55, 45, "dollars"+chr(9)+"percent", THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs)
					x_pos = x_pos + 195
					If x_pos = 415 Then
						x_pos = 25
						y_pos = y_pos + 20
					End If
				End If
			next
			If other_person_checkbox = checked Then
				Text x_pos, y_pos, 100, 10, "Other: " & other_person_name & " pays: "
				EditBox x_pos + 105, y_pos - 5, 50, 15, other_new_shel_total_amt
				DropListBox x_pos + 160, y_pos - 5, 55, 45, "dollars"+chr(9)+"percent", other_new_shel_total_amt_type
			End If

		    ' Text 25, 240, 60, 10, "Person ONE pays: "
		    ' EditBox 90, 235, 50, 15, Edit6
		    ' DropListBox 145, 235, 55, 45, "dollars"+chr(9)+"percent", List7
		    ' Text 220, 240, 70, 10, "Person THREE pays: "
		    ' EditBox 295, 235, 50, 15, Edit7
		    ' DropListBox 350, 235, 55, 45, "dollars"+chr(9)+"percent", List8
		    ' Text 25, 260, 65, 10, "Person TWO pays: "
		    ' EditBox 90, 255, 50, 15, Edit8
		    ' DropListBox 145, 255, 55, 45, "dollars"+chr(9)+"percent", List9
		    ' Text 225, 260, 65, 10, "Person FOUR pays: "
		    ' EditBox 295, 255, 50, 15, Edit9
		    ' DropListBox 350, 255, 55, 45, "dollars"+chr(9)+"percent", List10
		    PushButton 410, 260, 50, 10, "Enter", enter_shel_three_btn
		End If

		If shel_update_step > 3 Then
		    GroupBox 15, 280, 450, 55, "Is the housing expense verified?"
			x_pos = 25
			y_pos = 300

			If new_total_rent_amount <> "" AND new_total_rent_amount <> "0" Then
				Text x_pos, y_pos, 110, 10, "Total RENT of $" & new_total_rent_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_rent_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If

			If new_total_lot_rent_amount <> "" AND new_total_lot_rent_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 110, 10, "Total LOT RENT of $" & new_total_lot_rent_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_lot_rent_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_mortgage_amount <> "" AND new_total_mortgage_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 110, 10, "Total MORTGAGE of $" & new_total_mortgage_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_mortgage_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_insurance_amount <> "" AND new_total_insurance_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 120, 10, "Total INSURANCE of $" & new_total_insurance_amount & " verification:"
				DropListBox x_pos + 125, y_pos - 5, 80, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_insurance_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_tax_amount <> "" AND new_total_tax_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 110, 10, "Total TAXES of $" & new_total_tax_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_taxes_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_room_amount <> "" AND new_total_room_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 110, 10, "Total TOOM of $" & new_total_room_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_room_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_garage_amount <> "" AND new_total_garage_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 110, 10, "Total GARAGE of $" & new_total_garage_amount & " verification:"
				DropListBox x_pos + 115, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_garage_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
			If new_total_subsidy_amount <> "" AND new_total_subsidy_amount <> "0" Then
				' hold
				Text x_pos, y_pos, 120, 10, "Total SUBSIDY of $" & new_total_subsidy_amount & " verification:"
				DropListBox x_pos + 125, y_pos - 5, 80, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", new_subsidy_verif
				x_pos = x_pos + 200
				If x_pos = 425 Then
					x_pos = 25
					y_pos = y_pos + 15
				End If
			End If
		    ' Text 30, 300, 110, 10, "Total Rent of $XXXX verification:"
		    ' DropListBox 145, 295, 60, 45, "", List11
		    ' Text 235, 300, 110, 10, "Total Rent of $XXXX verification:"
		    ' DropListBox 350, 295, 60, 45, "", List12
		    ' Text 30, 315, 110, 10, "Total Rent of $XXXX verification:"
		    ' DropListBox 145, 310, 60, 45, "", List13
		    ' Text 235, 315, 110, 10, "Total Rent of $XXXX verification:"
		    ' DropListBox 350, 310, 60, 45, "", List14
		End If

		' Text 20, 145, 160, 10, "Have we received verification of this expense?"
		' DropListBox 180, 140, 45, 45, yes_no_list, shel_verif_received_yn

		' Text 40, 45, 120, 10, "Is the new expense amount shared?"
		' DropListBox 165, 40, 45, 45, yes_no_list, shel_shared_yn
		' Text 240, 45, 135, 10, "Is the new expense amount subsidized?"
		' DropListBox 380, 40, 45, 45, yes_no_list, shel_subsidized_yn

		' EditBox 105, 95, 45, 15, total_current_rent
		' DropListBox 155, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_rent_verif
		' EditBox 105, 115, 45, 15, total_current_lot_rent
		' DropListBox 155, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_lot_rent_verif
		' EditBox 105, 135, 45, 15, total_current_mortgage
		' DropListBox 155, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_mortgage_verif
		' EditBox 105, 155, 45, 15, total_current_insurance
		' DropListBox 155, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_insurance_verif
		' EditBox 105, 175, 45, 15, total_current_taxes
		' DropListBox 155, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_taxes_verif
		' EditBox 105, 195, 45, 15, total_current_room
		' DropListBox 155, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_room_verif
		' EditBox 105, 215, 45, 15, total_current_garage
		' DropListBox 155, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_garage_verif
		' EditBox 105, 235, 45, 15, total_current_subsidy
		' DropListBox 155, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", all_subsidy_verif
		'
		' Text 105, 75, 100, 10, "TOTAL Expense Amounts"
	    ' Text 105, 85, 30, 10, "Amount"
	    ' Text 160, 85, 20, 10, "Verif"
		' Text 80, 100, 20, 10, "Rent:"
	    ' Text 70, 120, 30, 10, "Lot Rent:"
	    ' Text 65, 140, 35, 10, "Mortgage:"
	    ' Text 65, 160, 40, 10, "Insurance:"
	    ' Text 75, 180, 25, 10, "Taxes:"
	    ' Text 75, 200, 25, 10, "Room:"
	    ' Text 75, 220, 30, 10, "Garage:"
	    ' Text 70, 240, 30, 10, "Subsidy:"


		' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
		' PushButton x_pos + 60, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	End If
	' If housing_questions_step = 4 Then
	' 	Text shel_det_x_pos + 5, 10, 60, 10, "SHEL DETAILS"
	'
	' 	Text 15, 25, 450, 10, "STEP 4 - SHEL Details"
	'
	' 	' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	' 	' PushButton x_pos + 60, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	' 	' PushButton x_pos + 120, 8, 60, 13, "SHEL UPDATE", housing_change_shel_update_btn
	' End If
	If housing_questions_step = 4 Then
		Text 420, 10, 60, 10, "REVIEW"

		Text 15, 25, 450, 10, "STEP 5 - REVIEW AND CONFIRM"

	End If

	Text 20, 350, 55, 10, "Additional Notes:"
	EditBox 80, 345, 385, 15, addr_or_shel_change_notes

	If housing_questions_step <> 1 Then PushButton overview_x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	If view_addr_update_dlg = True AND housing_questions_step <> 2 Then PushButton addr_up_x_pos, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	If view_shel_update_dlg = True AND housing_questions_step <> 3 Then PushButton shel_up_x_pos, 8, 60, 13, "SHEL UPDATE", housing_change_shel_update_btn
	' If view_shel_details_dlg = True AND housing_questions_step <> 4 Then PushButton shel_det_x_pos, 8, 60, 13, "SHEL DETAILS", housing_change_shel_details_btn
	If err_msg = "" AND housing_questions_step <> 4 Then PushButton 405, 8, 60, 13, "REVIEW", housing_change_review_btn

	If housing_questions_step <> 4 Then PushButton 390, 325, 70, 10, "CONTINUE", housing_change_continue_btn

end function

' function display_ADDR_information(update_addr, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)
'
' 	GroupBox 10, 35, 375, 95, "Residence Address"
' 	If update_addr = False Then
' 		Text 330, 35, 50, 10, addr_eff_date
' 		Text 70, 55, 305, 15, resi_street_full
' 		Text 70, 75, 105, 15, resi_city
' 		Text 205, 75, 110, 45, resi_state
' 		Text 340, 75, 35, 15, resi_zip
' 		Text 125, 95, 45, 45, addr_reservation
' 		Text 245, 85, 130, 15, reservation_name
' 		Text 125, 115, 45, 45, addr_homeless
' 		If addr_living_sit = "10 - Unknown" OR addr_living_sit = "Blank" Then
' 			DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
' 		Else
' 			Text 245, 115, 130, 45, addr_living_sit
' 		End If
' 		Text 70, 165, 305, 15, mail_street_full
' 		Text 70, 185, 105, 15, mail_city
' 		Text 205, 185, 110, 45, mail_state
' 		Text 340, 185, 35, 15, mail_zip
' 		Text 20, 240, 90, 15, phone_one
' 		Text 125, 240, 65, 45, type_one
' 		Text 20, 260, 90, 15, phone_two
' 		Text 125, 260, 65, 45, type_two
' 		Text 20, 280, 90, 15, phone_three
' 		Text 125, 280, 65, 45, type_three
' 		Text 325, 215, 50, 10, address_change_date
' 		Text 255, 245, 120, 10, resi_county
' 		Text 255, 280, 120, 10, addr_verif
' 		EditBox 10, 320, 375, 15, notes_on_address
' 		PushButton 290, 300, 95, 15, "Update Information", update_information_btn
' 	End If
' 	If update_addr = True Then
' 		EditBox 330, 30, 40, 15, addr_eff_date
' 		EditBox 70, 50, 305, 15, resi_street_full
' 		EditBox 70, 70, 105, 15, resi_city
' 		DropListBox 205, 70, 110, 45, ""+chr(9)+state_list, resi_state
' 		EditBox 340, 70, 35, 15, resi_zip
' 		DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", addr_reservation
' 		DropListBox 245, 90, 130, 15, "Select One..."+chr(9)+""+chr(9)+"BD - Bois Forte - Deer Creek"+chr(9)+"BN - Bois Forte - Nett Lake"+chr(9)+"BV - Bois Forte - Vermillion Lk"+chr(9)+"FL - Fond du Lac"+chr(9)+"GP - Grand Portage"+chr(9)+"LL - Leach Lake"+chr(9)+"LS - Lower Sioux"+chr(9)+"ML - Mille Lacs"+chr(9)+"PL - Prairie Island Community"+chr(9)+"RL - Red Lake"+chr(9)+"SM - Shakopee Mdewakanton"+chr(9)+"US - Upper Sioux"+chr(9)+"WE - White Earth", reservation_name
' 		DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", addr_homeless
' 		DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
' 		EditBox 70, 160, 305, 15, mail_street_full
' 		EditBox 70, 180, 105, 15, mail_city
' 		DropListBox 205, 180, 110, 45, ""+chr(9)+state_list, mail_state
' 		EditBox 340, 180, 35, 15, mail_zip
' 		EditBox 20, 240, 90, 15, phone_one
' 		DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_one
' 		EditBox 20, 260, 90, 15, phone_two
' 		DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_two
' 		EditBox 20, 280, 90, 15, phone_three
' 		DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_three
' 		EditBox 325, 210, 50, 15, address_change_date
' 		ComboBox 255, 245, 120, 45, county_list_smalll+chr(9)+resi_county, resi_county
' 		DropListBox 255, 280, 120, 45, "SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed"+chr(9)+"Blank", addr_verif
' 		EditBox 10, 320, 375, 15, notes_on_address
' 		PushButton 290, 300, 95, 15, "Save Information", save_information_btn
' 	End If
'
' 	PushButton 325, 145, 50, 10, "CLEAR", clear_mail_addr_btn
' 	PushButton 205, 240, 35, 10, "CLEAR", clear_phone_one_btn
' 	PushButton 205, 260, 35, 10, "CLEAR", clear_phone_two_btn
' 	PushButton 205, 280, 35, 10, "CLEAR", clear_phone_three_btn
' 	' Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
' 	Text 250, 35, 80, 10, "ADDR effective date:"
' 	Text 20, 55, 45, 10, "House/Street"
' 	Text 45, 75, 20, 10, "City"
' 	Text 185, 75, 20, 10, "State"
' 	Text 325, 75, 15, 10, "Zip"
' 	Text 20, 95, 100, 10, "Do you live on a Reservation?"
' 	Text 180, 95, 60, 10, "If yes, which one?"
' 	Text 30, 115, 90, 10, "Client Indicates Homeless:"
' 	Text 185, 115, 60, 10, "Living Situation?"
' 	GroupBox 10, 135, 375, 70, "Mailing Address"
' 	Text 20, 165, 45, 10, "House/Street"
' 	Text 45, 185, 20, 10, "City"
' 	Text 185, 185, 20, 10, "State"
' 	Text 325, 185, 15, 10, "Zip"
' 	GroupBox 10, 210, 235, 90, "Phone Number"
' 	Text 20, 225, 50, 10, "Number"
' 	Text 125, 225, 25, 10, "Type"
' 	Text 255, 215, 60, 10, "Date of Change:"
' 	Text 255, 235, 75, 10, "County of Residence:"
' 	Text 255, 270, 75, 10, "ADDR Verification:"
' 	Text 10, 310, 75, 10, "Additional Notes:"
' end function
'
' function display_SHEL_information(update_shel, show_totals, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, update_information_btn, save_information_btn, const_memb_buttons, clear_all_btn, view_total_shel_btn, update_household_percent_button)
'
' 	Text 10, 10, 360, 10, "Review the Shelter informaiton known with the client. If it needs updating, press this button to make changes:"
' 	y_pos = 25
' 	For the_member = 0 to UBound(SHEL_ARRAY, 2)
' 		If the_member = selection Then
' 			Text 416, y_pos + 2, 60, 10, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member)
' 			y_pos = y_pos + 15
' 		Else
' 			PushButton 400, y_pos, 75, 13, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member), SHEL_ARRAY(const_memb_buttons, the_member)
' 			y_pos = y_pos + 15
' 		End If
' 	Next
' 	' MsgBox "In DISPLAY" & vbCr & vbCr & "Show totals - " & show_totals
' 	If show_totals = True Then
' 		Text 415, 223, 65, 10, "TOTAL SHEL"
'
' 		If update_shel = True Then
' 			EditBox 105, 25, 165, 15, total_paid_to
' 			EditBox 125, 40, 20, 15, total_paid_by_household
' 			EditBox 125, 55, 20, 15, total_paid_by_others
' 			EditBox 105, 95, 45, 15, total_current_rent
' 			EditBox 105, 115, 45, 15, total_current_lot_rent
' 			EditBox 105, 135, 45, 15, total_current_mortgage
' 			EditBox 105, 155, 45, 15, total_current_insurance
' 			EditBox 105, 175, 45, 15, total_current_taxes
' 			EditBox 105, 195, 45, 15, total_current_room
' 			EditBox 105, 215, 45, 15, total_current_garage
' 			EditBox 105, 235, 45, 15, total_current_subsidy
' 			PushButton 400, 235, 75, 15, "Save Information", save_information_btn
' 		End If
' 		If update_shel = False Then
' 			Text 105, 30, 165, 10, total_paid_to
' 			Text 125, 45, 20, 10, total_paid_by_household
' 			Text 125, 60, 20, 10, total_paid_by_others
' 			Text 105, 100, 45, 10, total_current_rent
' 			Text 105, 120, 45, 10, total_current_lot_rent
' 			Text 105, 140, 45, 10, total_current_mortgage
' 			Text 105, 160, 45, 10, total_current_insurance
' 			Text 105, 180, 45, 10, total_current_taxes
' 			Text 105, 200, 45, 10, total_current_room
' 			Text 105, 220, 45, 10, total_current_garage
' 			Text 105, 240, 45, 10, total_current_subsidy
' 			PushButton 400, 235, 75, 15, "Update Information", update_information_btn
' 		End If
' 		Text 15, 30, 90, 10, "Housing Expense Paid to"
' 		Text 15, 45, 100, 10, "Expense Paid by Household"
' 		Text 145, 45, 50, 10, "% (percent)"
' 		PushButton 210, 41, 125, 13, "MANAGE HOUSEHOLD PERCENT", update_household_percent_button
' 		Text 15, 60, 100, 10, "Expense Paid by Someone Else"
' 		Text 145, 60, 50, 10, "% (percent)"
' 		Text 105, 75, 60, 20, "Current Total Amount"
' 		Text 80, 100, 20, 10, "Rent:"
' 	    Text 70, 120, 30, 10, "Lot Rent:"
' 	    Text 65, 140, 35, 10, "Mortgage:"
' 	    Text 65, 160, 40, 10, "Insurance:"
' 	    Text 75, 180, 25, 10, "Taxes:"
' 	    Text 75, 200, 25, 10, "Room:"
' 	    Text 75, 220, 30, 10, "Garage:"
' 	    Text 70, 240, 30, 10, "Subsidy:"
'
' 	Else
' 		PushButton 400, 220, 75, 15, "TOTAL SHEL", view_total_shel_btn
'
' 		If update_shel = True Then
' 			EditBox 105, 25, 165, 15, SHEL_ARRAY(const_paid_to, selection)
' 			DropListBox 165, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_hud_sub_yn, selection)
' 			DropListBox 310, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_shared_yn, selection)
' 			EditBox 105, 95, 45, 15, SHEL_ARRAY(const_rent_retro_amt, selection)
' 			DropListBox 155, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_retro_verif, selection)
' 			EditBox 255, 95, 45, 15, SHEL_ARRAY(const_rent_prosp_amt, selection)
' 			DropListBox 305, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_prosp_verif, selection)
' 			EditBox 105, 115, 45, 15, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
' 			DropListBox 155, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_retro_verif, selection)
' 			EditBox 255, 115, 45, 15, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
' 			DropListBox 305, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
' 			EditBox 105, 135, 45, 15, SHEL_ARRAY(const_mortgage_retro_amt, selection)
' 			DropListBox 155, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_retro_verif, selection)
' 			EditBox 255, 135, 45, 15, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
' 			DropListBox 305, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_prosp_verif, selection)
' 			EditBox 105, 155, 45, 15, SHEL_ARRAY(const_insurance_retro_amt, selection)
' 			DropListBox 155, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_retro_verif, selection)
' 			EditBox 255, 155, 45, 15, SHEL_ARRAY(const_insurance_prosp_amt, selection)
' 			DropListBox 305, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_prosp_verif, selection)
' 			EditBox 105, 175, 45, 15, SHEL_ARRAY(const_tax_retro_amt, selection)
' 			DropListBox 155, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_retro_verif, selection)
' 			EditBox 255, 175, 45, 15, SHEL_ARRAY(const_tax_prosp_amt, selection)
' 			DropListBox 305, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_prosp_verif, selection)
' 			EditBox 105, 195, 45, 15, SHEL_ARRAY(const_room_retro_amt, selection)
' 			DropListBox 155, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_retro_verif, selection)
' 			EditBox 255, 195, 45, 15, SHEL_ARRAY(const_room_prosp_amt, selection)
' 			DropListBox 305, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_prosp_verif, selection)
' 			EditBox 105, 215, 45, 15, SHEL_ARRAY(const_garage_retro_amt, selection)
' 			DropListBox 155, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_retro_verif, selection)
' 			EditBox 255, 215, 45, 15, SHEL_ARRAY(const_garage_prosp_amt, selection)
' 			DropListBox 305, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_prosp_verif, selection)
' 			EditBox 105, 235, 45, 15, SHEL_ARRAY(const_subsidy_retro_amt, selection)
' 			DropListBox 155, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_retro_verif, selection)
' 			EditBox 255, 235, 45, 15, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
' 			DropListBox 305, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_prosp_verif, selection)
' 			PushButton 400, 235, 75, 15, "Save Information", save_information_btn
' 		End If
' 		If update_shel = False Then
' 			Text 105, 30, 165, 10, SHEL_ARRAY(const_paid_to, selection)
' 			Text 165, 50, 40, 10, SHEL_ARRAY(const_hud_sub_yn, selection)
' 			Text 310, 50, 40, 10, SHEL_ARRAY(const_shared_yn, selection)
' 			Text 105, 100, 45, 10, SHEL_ARRAY(const_rent_retro_amt, selection)
' 			Text 160, 100, 70, 10, SHEL_ARRAY(const_rent_retro_verif, selection)
' 			Text 255, 100, 45, 10, SHEL_ARRAY(const_rent_prosp_amt, selection)
' 			Text 310, 100, 70, 10, SHEL_ARRAY(const_rent_prosp_verif, selection)
' 			Text 105, 120, 45, 10, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
' 			Text 160, 120, 70, 10, SHEL_ARRAY(const_lot_rent_retro_verif, selection)
' 			Text 255, 120, 45, 10, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
' 			Text 310, 120, 70, 10, SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
' 			Text 105, 140, 45, 10, SHEL_ARRAY(const_mortgage_retro_amt, selection)
' 			Text 160, 140, 70, 10, SHEL_ARRAY(const_mortgage_retro_verif, selection)
' 			Text 255, 140, 45, 10, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
' 			Text 310, 140, 70, 10, SHEL_ARRAY(const_mortgage_prosp_verif, selection)
' 			Text 105, 160, 45, 10, SHEL_ARRAY(const_insurance_retro_amt, selection)
' 			Text 160, 160, 70, 10, SHEL_ARRAY(const_insurance_retro_verif, selection)
' 			Text 255, 160, 45, 10, SHEL_ARRAY(const_insurance_prosp_amt, selection)
' 			Text 310, 160, 70, 10, SHEL_ARRAY(const_insurance_prosp_verif, selection)
' 			Text 105, 180, 45, 10, SHEL_ARRAY(const_tax_retro_amt, selection)
' 			Text 160, 180, 70, 10, SHEL_ARRAY(const_tax_retro_verif, selection)
' 			Text 255, 180, 45, 10, SHEL_ARRAY(const_tax_prosp_amt, selection)
' 			Text 310, 180, 70, 10, SHEL_ARRAY(const_tax_prosp_verif, selection)
' 			Text 105, 200, 45, 10, SHEL_ARRAY(const_room_retro_amt, selection)
' 			Text 160, 200, 70, 10, SHEL_ARRAY(const_room_retro_verif, selection)
' 			Text 255, 200, 45, 10, SHEL_ARRAY(const_room_prosp_amt, selection)
' 			Text 310, 200, 70, 10, SHEL_ARRAY(const_room_prosp_verif, selection)
' 			Text 105, 220, 45, 10, SHEL_ARRAY(const_garage_retro_amt, selection)
' 			Text 160, 220, 70, 10, SHEL_ARRAY(const_garage_retro_verif, selection)
' 			Text 255, 220, 45, 10, SHEL_ARRAY(const_garage_prosp_amt, selection)
' 			Text 310, 220, 70, 10, SHEL_ARRAY(const_garage_prosp_verif, selection)
' 			Text 105, 240, 45, 10, SHEL_ARRAY(const_subsidy_retro_amt, selection)
' 			Text 160, 240, 70, 10, SHEL_ARRAY(const_subsidy_retro_verif, selection)
' 			Text 255, 240, 45, 10, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
' 			Text 310, 240, 70, 10, SHEL_ARRAY(const_subsidy_prosp_verif, selection)
' 			PushButton 400, 235, 75, 15, "Update Information", update_information_btn
' 		End If
'
' 		PushButton 325, 30, 70, 13, "CLEAR ALL", clear_all_btn
' 	    Text 15, 30, 90, 10, "Housing Expense Paid to"
' 		Text 105, 50, 60, 10, "HUD Subsidized"
' 	    Text 225, 50, 85, 10, "Housing Expense Shared"
' 	    GroupBox 15, 65, 380, 190, "Housing Expense Amounts"
' 	    Text 105, 75, 65, 10, "Retrospective"
' 	    Text 255, 75, 65, 10, "Prospective"
' 	    Text 105, 85, 30, 10, "Amount"
' 	    Text 255, 85, 25, 10, "Amount"
' 	    Text 160, 85, 20, 10, "Verif"
' 	    Text 310, 85, 20, 10, "Verif"
' 		Text 80, 100, 20, 10, "Rent:"
' 	    Text 70, 120, 30, 10, "Lot Rent:"
' 	    Text 65, 140, 35, 10, "Mortgage:"
' 	    Text 65, 160, 40, 10, "Insurance:"
' 	    Text 75, 180, 25, 10, "Taxes:"
' 	    Text 75, 200, 25, 10, "Room:"
' 	    Text 75, 220, 30, 10, "Garage:"
' 	    Text 70, 240, 30, 10, "Subsidy:"
'
' 	End If
'
'
'
' 	'CAF Questions'
' 	' Text 20, 270, 125, 10, "Rent (include mobild home lot rental)"
'     ' DropListBox 145, 265, 40, 45, "caf_answer_droplist", q14_rent_caf_answer
'     ' EditBox 190, 265, 35, 15, q14_rent_caf_response
'     ' Text 20, 285, 125, 10, "Mortgage/Contract for Deed Payment"
'     ' DropListBox 145, 280, 40, 45, "caf_answer_droplist", q14_mort_caf_answer
'     ' EditBox 190, 280, 35, 15, q14_mort_caf_response
'     ' Text 20, 300, 125, 10, "Homeowner's Insurance"
'     ' DropListBox 145, 295, 40, 45, "caf_answer_droplist", q14_ins_caf_answer
'     ' EditBox 190, 295, 35, 15, q14_ins_caf_response
'     ' Text 20, 315, 125, 10, "Real Estate Taxes"
'     ' DropListBox 145, 310, 40, 45, "caf_answer_droplist", q14_tax_caf_answer
'     ' EditBox 190, 310, 35, 15, q14_tax_caf_response
'     ' Text 240, 270, 105, 10, "Rental or Secontion 8 Subsidy"
'     ' DropListBox 345, 265, 40, 45, "caf_answer_droplist", q14_subs_caf_answer
'     ' EditBox 390, 265, 35, 15, q14_subs_caf_response
'     ' Text 240, 285, 100, 10, "Association Fees"
'     ' DropListBox 345, 280, 40, 45, "caf_answer_droplist", q14_fees_caf_answer
'     ' EditBox 390, 280, 35, 15, q14_fees_caf_response
'     ' Text 240, 300, 95, 10, "Room and/or Board"
'     ' DropListBox 345, 295, 40, 45, "caf_answer_droplist", q14_room_caf_answer
'     ' EditBox 390, 295, 35, 15, q14_room_caf_response
'     ' Text 240, 315, 105, 20, "CONFIM - Do you get help paying rent?"
'     ' DropListBox 345, 310, 40, 45, "caf_answer_droplist", q14_confirm_subsidy
'     ' EditBox 390, 310, 35, 15, q14_confirm_subsidy_amount
' end function

' function display_HEST_information(update_hest, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
'
' 	If update_hest = False Then
' 		Text 75, 30, 145, 10, all_persons_paying
' 	    Text 75, 50, 50, 10, choice_date
' 	    Text 125, 70, 50, 10, actual_initial_exp
' 	    Text 70, 125, 40, 10, retro_heat_ac_yn
' 	    Text 115, 125, 20, 10, retro_heat_ac_units
' 	    Text 150, 125, 45, 10, retro_heat_ac_amt
' 	    Text 240, 125, 40, 10, prosp_heat_ac_yn
' 	    Text 285, 125, 20, 10, prosp_heat_ac_units
' 	    Text 320, 125, 45, 10, prosp_heat_ac_amt
' 	    Text 70, 145, 40, 10, retro_electric_yn
' 	    Text 115, 145, 20, 10, retro_electric_units
' 	    Text 150, 145, 45, 10, retro_electric_amt
' 	    Text 240, 145, 40, 10, prosp_electric_yn
' 	    Text 285, 145, 20, 10, prosp_electric_units
' 	    Text 320, 145, 45, 10, prosp_electric_amt
' 	    Text 70, 165, 40, 10, retro_phone_yn
' 	    Text 115, 165, 20, 10, retro_phone_units
' 	    Text 150, 165, 45, 10, retro_phone_amt
' 	    Text 240, 165, 40, 10, prosp_phone_yn
' 	    Text 285, 165, 20, 10, prosp_phone_units
' 	    Text 320, 165, 45, 10, prosp_phone_amt
' 		Text 55, 185, 150, 10, "Total Counted Utility Expense: $" & total_utility_expense
'
' 		PushButton 290, 185, 95, 15, "Update Information", update_information_btn
' 	End If
' 	If update_hest = True Then
' 		EditBox 75, 25, 145, 15, all_persons_paying
' 	    EditBox 75, 45, 50, 15, choice_date
' 	    EditBox 125, 65, 50, 15, actual_initial_exp
' 	    DropListBox 65, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_heat_ac_yn
' 	    ' EditBox 110, 120, 20, 15, retro_heat_ac_units
' 	    ' EditBox 150, 120, 45, 15, retro_heat_ac_amt
' 	    DropListBox 235, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_heat_ac_yn
' 	    ' EditBox 280, 120, 20, 15, prosp_heat_ac_units
' 	    ' EditBox 320, 120, 45, 15, prosp_heat_ac_amt
' 	    DropListBox 65, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_electric_yn
' 	    ' EditBox 110, 140, 20, 15, retro_electric_units
' 	    ' EditBox 150, 140, 45, 15, retro_electric_amt
' 	    DropListBox 235, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_electric_yn
' 	    ' EditBox 280, 140, 20, 15, prosp_electric_units
' 	    ' EditBox 320, 140, 45, 15, prosp_electric_amt
' 	    DropListBox 65, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_phone_yn
' 	    ' EditBox 110, 160, 20, 15, retro_phone_units
' 	    ' EditBox 150, 160, 45, 15, retro_phone_amt
' 	    DropListBox 235, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_phone_yn
' 	    ' EditBox 280, 160, 20, 15, prosp_phone_units
' 	    ' EditBox 320, 160, 45, 15, prosp_phone_amt
' 		' ComboBox 255, 255, 120, 45, county_list+chr(9)+resi_addr_county, resi_addr_county
' 		PushButton 290, 185, 95, 15, "Save Information", save_information_btn
' 	End If
'
'
' 	Text 10, 10, 360, 10, "Review the Utility Information"
'     Text 15, 30, 60, 10, "Persons Paying:"
'     Text 15, 50, 55, 10, "FS Choice Date:"
'     Text 15, 70, 110, 10, "Actual Expense In Initial Month: $ "
'     Text 20, 125, 30, 10, "Heat/Air:"
'     Text 20, 145, 30, 10, "Electric:"
'     Text 25, 165, 25, 10, "Phone:"
'     GroupBox 55, 85, 150, 95, "Retrospective"
'     Text 65, 105, 20, 10, "(Y/N)"
'     Text 110, 100, 20, 20, "#/FS Units"
'     Text 150, 105, 30, 10, "Amount"
'     GroupBox 225, 85, 150, 95, "Prospective"
'     Text 235, 105, 20, 10, "(Y/N)"
'     Text 280, 100, 20, 20, "#/FS Units"
'     Text 320, 105, 25, 10, "Amount"
'
' 	' GroupBox 20, 150, 455, grp_len, "Already Known Shelter Expenses - Added or listed in MAXIS"
' 	' ' Text 30, 165, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
' 	' ' Text 30, 180, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
' 	' PushButton 350, y_pos, 125, 10, "Update Shelter Expense Information", update_shel_btn
' 	' y_pos = y_pos + 15
' 	' Text 5, y_pos, 310, 10, "^^4 - Enter the answers listed on the actual CAF form for Q15 into the 'Answer on the CAF' field."
' 	' Text 20, y_pos + 10, 295, 10, "Q. 15. Does your household have the following utility expenses any time during the year?"
' 	' y_pos = y_pos + 30
' 	' Text 20, y_pos, 85, 10, "Heating/Air Conditioning"
' 	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_ac_caf_answer
' 	' Text 180, y_pos, 85, 10, "Electricity"
' 	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_answer
' 	' Text 345, y_pos, 85, 10, "Cooking Fuel"
' 	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_answer
' 	' y_pos = y_pos + 15
' 	' Text 20, y_pos, 85, 10, "Water and Sewer"
' 	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_answer
' 	' Text 180, y_pos, 85, 10, "Garbage Removal"
' 	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_answer
' 	' Text 345, y_pos, 85, 10, "Phone/Cell Phone"
' 	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_answer
' 	' y_pos = y_pos + 15
' 	' Text 75, y_pos, 355, 10, "Did anyone in the household receive Energy Assistance (LIHEAP) of more than $20 in the past 12 months?"
' 	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_liheap_caf_answer
' 	' y_pos = y_pos + 15
' 	'
' 	' Text 5, y_pos, 270, 10, "^^5 - ASK - 'Does anyone in the household pay ...'  RECORD the verbal responses"
' 	' y_pos = y_pos + 20
' 	' Text 20, y_pos, 85, 10, "Heating"
' 	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_caf_response
' 	' Text 180, y_pos, 85, 10, "Electricity"
' 	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_response
' 	' Text 345, y_pos, 85, 10, "Cooking Fuel"
' 	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_response
' 	' y_pos = y_pos + 15
' 	' Text 20, y_pos, 85, 10, "Air Conditioning"
' 	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ac_caf_response
' 	' Text 180, y_pos, 85, 10, "Garbage Removal"
' 	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_response
' 	' Text 345, y_pos, 85, 10, "Phone/Cell Phone"
' 	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_response
' 	' y_pos = y_pos + 15
' 	' Text 20, y_pos, 85, 10, "Water and Sewer"
' 	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_response
' 	' Text 170, y_pos + 5, 265, 10, "Did your household receive any help in paying for your energy or power bills?"
' 	' DropListBox 435, y_pos, 40, 45, caf_answer_droplist, q15_liheap_caf_response
' 	' y_pos = y_pos + 15
' 	' PushButton 20, y_pos, 130, 10, "Utilities are Complicated", utility_detail_btn
' end function

' function navigate_ADDR_buttons(update_addr, err_var, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, mail_street_full, mail_city, mail_state, mail_zip, phone_one, phone_two, phone_three, type_one, type_two, type_three)
' 	If ButtonPressed = update_information_btn Then
' 		update_addr = TRUE
' 		' update_attempted = True
' 	ElseIf ButtonPressed = save_information_btn Then
' 		update_addr = FALSE
' 	Else
' 		update_addr = FALSE
' 	End If
'
' 	If ButtonPressed = clear_mail_addr_btn Then
' 		mail_street_full = ""
' 		mail_city = ""
' 		mail_state = ""
' 		mail_zip = ""
' 	End If
' 	If ButtonPressed = clear_phone_one_btn Then
' 		phone_one = ""
' 		type_one = "Select One..."
' 	End If
' 	If ButtonPressed = clear_phone_two_btn Then
' 		phone_two = ""
' 		type_two = "Select One..."
' 	End If
' 	If ButtonPressed = clear_phone_three_btn Then
' 		phone_three = ""
' 		type_three = "Select One..."
' 	End If
' end function

' function navigate_SHEL_buttons(update_shel, err_var, update_attempted, update_information_btn, save_information_btn, SHEL_ARRAY, const_memb_buttons, const_shel_exists, const_attempt_update, selection)

' function navigate_SHEL_buttons(update_shel, show_totals, err_var, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, update_information_btn, save_information_btn, const_memb_buttons, const_attempt_update, clear_all_btn, view_total_shel_btn)
'
' 	If ButtonPressed = update_information_btn Then
' 		update_shel = TRUE
' 		update_attempted = True
' 		' MsgBox "In UPDATE button" & vbCr & vbCr & "Show totals - " & show_totals
' 	ElseIf ButtonPressed = save_information_btn Then
' 		update_shel = FALSE
' 	Else
' 		update_shel = FALSE
' 	End If
'
' 	If selection <> "" Then
' 		'REVIEWING THE INFORMATION IN THE ARRAY TO DETERMINE IF IT IS BLANK
' 		all_shel_details_blank = True
'
' 		If Trim(SHEL_ARRAY(const_paid_to, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_hud_sub_yn, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_shared_yn, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_lot_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_lot_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_lot_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_lot_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_mortgage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_mortgage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_mortgage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_mortgage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_insurance_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_insurance_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_insurance_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_insurance_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_tax_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_tax_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_tax_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_tax_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_room_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_room_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_room_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_room_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_garage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_garage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_garage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_garage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_subsidy_retro_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_subsidy_retro_verif, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_subsidy_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
' 		If Trim(SHEL_ARRAY(const_subsidy_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
'
' 		If all_shel_details_blank = True Then SHEL_ARRAY(const_shel_exists, selection) = False
'
' 		If ButtonPressed = clear_all_btn Then
' 			SHEL_ARRAY(const_paid_to, selection) = ""
' 			SHEL_ARRAY(const_hud_sub_yn, selection) = ""
' 			SHEL_ARRAY(const_shared_yn, selection) = ""
' 			SHEL_ARRAY(const_rent_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_rent_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_rent_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_rent_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_lot_rent_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_lot_rent_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_lot_rent_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_mortgage_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_mortgage_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_mortgage_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_insurance_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_insurance_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_insurance_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_insurance_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_tax_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_tax_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_tax_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_tax_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_room_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_room_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_room_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_room_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_garage_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_garage_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_garage_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_garage_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_subsidy_retro_amt, selection) = ""
' 			SHEL_ARRAY(const_subsidy_retro_verif, selection) = ""
' 			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = ""
' 			SHEL_ARRAY(const_subsidy_prosp_verif, selection) = ""
' 			SHEL_ARRAY(const_shel_exists, selection) = False
' 		End If
' 	End If
'
' 	For memb_btn = 0 to UBound(SHEL_ARRAY, 2)
' 		If ButtonPressed = SHEL_ARRAY(const_memb_buttons, memb_btn) Then
' 			selection = memb_btn
' 			show_totals = False
' 		End If
' 	Next
' 	If selection <> "" Then
' 		If SHEL_ARRAY(const_shel_exists, selection) = False Then update_shel = True
' 		If update_shel = True Then
' 			SHEL_ARRAY(const_attempt_update, selection) = True
' 			update_attempted = True
'
' 			SHEL_ARRAY(const_rent_prosp_amt, selection) = SHEL_ARRAY(const_rent_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = SHEL_ARRAY(const_lot_rent_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = SHEL_ARRAY(const_mortgage_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_insurance_prosp_amt, selection) = SHEL_ARRAY(const_insurance_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_tax_prosp_amt, selection) = SHEL_ARRAY(const_tax_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_room_prosp_amt, selection) = SHEL_ARRAY(const_room_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_garage_prosp_amt, selection) = SHEL_ARRAY(const_garage_prosp_amt, selection) & ""
' 			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = SHEL_ARRAY(const_subsidy_prosp_amt, selection) & ""
' 		End If
' 	End If
' 	If ButtonPressed = view_total_shel_btn Then
' 		show_totals = True
' 		selection = ""
' 	End If
' 	If show_totals = True and update_shel = True Then
' 		total_paid_by_household = total_paid_by_household & ""
' 		total_paid_by_others = total_paid_by_others & ""
' 		total_current_rent = total_current_rent & ""
' 		total_current_lot_rent = total_current_lot_rent & ""
' 		total_current_mortgage = total_current_mortgage & ""
' 		total_current_insurance = total_current_insurance & ""
' 		total_current_taxes = total_current_taxes & ""
' 		total_current_room = total_current_room & ""
' 		total_current_garage = total_current_garage & ""
' 		total_current_subsidy = total_current_subsidy & ""
' 	End If
' 	' MsgBox "End NAVIGATE" & vbCr & vbCr & "Show totals - " & show_totals
' end function

' function navigate_HEST_buttons(update_hest, err_var, update_attempted, update_information_btn, save_information_btn, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, date_to_use_for_HEST_standards)
' 	Call hest_standards(heat_AC_amt, electric_amt, phone_amt, date_to_use_for_HEST_standards)
' 	If ButtonPressed = update_information_btn Then
' 		update_hest = TRUE
' 		update_attempted = True
'
' 		retro_heat_ac_amt = retro_heat_ac_amt & ""
' 		retro_electric_amt = retro_electric_amt & ""
' 		retro_phone_amt = retro_phone_amt & ""
' 		prosp_heat_ac_amt = prosp_heat_ac_amt & ""
' 		prosp_electric_amt = prosp_electric_amt & ""
' 		prosp_phone_amt = prosp_phone_amt & ""
'
' 	ElseIf ButtonPressed = save_information_btn Then
' 		update_hest = FALSE
'
' 		retro_heat_ac_amt = 0
' 		retro_heat_ac_units = ""
' 		retro_electric_amt = 0
' 		retro_electric_units = ""
' 		retro_phone_amt = 0
' 		retro_phone_units = ""
' 		prosp_heat_ac_amt = 0
' 		prosp_heat_ac_units = ""
' 		prosp_electric_amt = 0
' 		prosp_electric_units = ""
' 		prosp_phone_amt = 0
' 		prosp_phone_units = ""
'
' 		If retro_heat_ac_yn = "Y" Then
' 			retro_heat_ac_amt = heat_AC_amt
' 			retro_heat_ac_units = "01"
' 		End If
' 		If retro_electric_yn = "Y" Then
' 			retro_electric_amt = electric_amt
' 			retro_electric_units = "01"
' 		End If
' 		If retro_phone_yn = "Y" Then
' 			retro_phone_amt = phone_amt
' 			retro_phone_units = "01"
' 		End If
' 		If prosp_heat_ac_yn = "Y" Then
' 			prosp_heat_ac_amt = heat_AC_amt
' 			prosp_heat_ac_units = "01"
' 		End If
' 		If prosp_electric_yn = "Y" Then
' 			prosp_electric_amt = electric_amt
' 			prosp_electric_units = "01"
' 		End If
' 		If prosp_phone_yn = "Y" Then
' 			prosp_phone_amt = phone_amt
' 			prosp_phone_units = "01"
' 		End If
'
' 		total_utility_expense = 0
' 		If prosp_heat_ac_yn = "Y" Then
' 			total_utility_expense =  heat_AC_amt
' 		ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
' 			total_utility_expense =  electric_amt + phone_amt
' 		ElseIf prosp_electric_yn = "Y" Then
' 			total_utility_expense =  electric_amt
' 		Elseif prosp_phone_yn = "Y" Then
' 			total_utility_expense =  phone_amt
' 		End If
'
' 	Else
' 		update_hest = FALSE
' 	End If
' end function

function navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, THE_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_total_shelter_expense_amount, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn, update_shel_button, addr_update_needed, shel_update_needed, hest_update_needed)

	' view_addr_update_dlg
	' view_shel_update_dlg
	' view_shel_details_dlg
	start_on_shel_questions = True
	If housing_questions_step <> 3 Then start_on_shel_questions = False

	If housing_questions_step = 3 Then

		If (what_is_the_living_arrangement = "Apartment or Townhouse" OR what_is_the_living_arrangement = "House") AND unit_owned = "Select One..." Then err_msg = err_msg & vbCr & "* For Apartment, House, or Mobile Home, you must select if the unit is owned or not to continue."
		If new_rent_subsidy_yn = "Select One..." Then err_msg = err_msg & vbCr & "* You must indiccate if the rent is subsidized."
		If trim(new_renter_insurance_amount) <> "" AND new_renter_insurance_amount <> "0" and new_renters_insurance_required_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have indicated a renters insurance amount, you must indicate if the if the insurance is required by the lease."
		If trim(new_total_garage_amount) <> "" AND new_total_garage_amount <> "0" and new_garage_rent_required_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have indicated a garage expense, you must indicate if the garage rent is required by the leease."
		If trim(new_total_mortgage_amount) <> "" AND new_total_mortgage_amount <> "0" and new_mortgage_have_escrow_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since you have entered a mortgage amount, you must indicate if there is an escrow with that mortgage payment (if taxes and insurance are included in the expense amount)."
		If trim(new_total_mortgage_amount) <> "" AND new_total_mortgage_amount <> "0" and new_excess_insurance_yn = "Select One..." Then err_msg = err_msg & vbCr & "* Since this has a mortgage expensee, you must indicate if the insurance paid has excess ccoverage."
		If new_mortgage_have_escrow_yn = "No" AND trim(new_morgage_insurance_amount) = "" Then err_msg = err_msg & vbCr & "* Since you indicated the mortgage does not have an escrow, you must indicate what the insurance amount is."
		If new_mortgage_have_escrow_yn = "No" AND trim(new_total_tax_amount) = "" Then err_msg = err_msg & vbCr & "* Since you indicated the mortgage does not have an escrow, you must indicate what the tax amount is."
		' If what_is_the_living_arrangement = "Trailer Home/Mobile Home" Then err_msg = err_msg & vbCr & "* "
		If trim(new_total_room_amount) <> "" AND new_total_room_amount <> "0" AND new_room_payment_frequency = "Select One..." Then err_msg = err_msg & vbCr & "* You must indicate the frequency of the room expense payment."

		total_current_rent = trim(total_current_rent)
		If total_current_rent = "" Then total_current_rent = 0
		total_current_rent = total_current_rent * 1
		total_current_lot_rent = trim(total_current_lot_rent)
		If total_current_lot_rent = "" Then total_current_lot_rent = 0
		total_current_lot_rent = total_current_lot_rent * 1
		total_current_garage = trim(total_current_garage)
		If total_current_garage = "" Then total_current_garage = 0
		total_current_garage = total_current_garage * 1
		total_current_insurance = trim(total_current_insurance)
		If total_current_insurance = "" Then total_current_insurance = 0
		total_current_insurance = total_current_insurance * 1
		total_current_taxes = trim(total_current_taxes)
		If total_current_taxes = "" Then total_current_taxes = 0
		total_current_taxes = total_current_taxes * 1
		total_current_room = trim(total_current_room)
		If total_current_room = "" Then total_current_room = 0
		total_current_room = total_current_room * 1
		total_current_mortgage = trim(total_current_mortgage)
		If total_current_mortgage = "" Then total_current_mortgage = 0
		total_current_mortgage = total_current_mortgage * 1
		total_current_subsidy = trim(total_current_subsidy)
		If total_current_subsidy = "" Then total_current_subsidy = 0
		total_current_subsidy = total_current_subsidy * 1
		' all_rent_verif,
		' all_lot_rent_verif,
		' all_mortgage_verif,
		' all_insurance_verif,
		' all_taxes_verif,
		' all_room_verif,
		' all_garage_verif,
		' all_subsidy_verif,
		' total_shel_original_information)

		' new_total_rent_amount
		' new_renter_insurance_amount
		' new_total_garage_amount
		' new_total_mortgage_amount
		' new_morgage_insurance_amount
		' new_total_tax_amount
		' new_total_lot_rent_amount
		' new_total_room_amount
		' new_vehicle_insurance_amount
		' new_total_insurance_amount

		If new_total_rent_amount = "" Then new_total_rent_amount = 0
		new_total_rent_amount = new_total_rent_amount * 1
		If new_renter_insurance_amount = "" Then new_renter_insurance_amount = 0
		new_renter_insurance_amount = new_renter_insurance_amount * 1
		If new_total_garage_amount = "" Then new_total_garage_amount = 0
		new_total_garage_amount = new_total_garage_amount * 1
		If new_total_mortgage_amount = "" Then new_total_mortgage_amount = 0
		new_total_mortgage_amount = new_total_mortgage_amount * 1
		If new_morgage_insurance_amount = "" Then new_morgage_insurance_amount = 0
		new_morgage_insurance_amount = new_morgage_insurance_amount * 1
		If new_total_tax_amount = "" Then new_total_tax_amount = 0
		new_total_tax_amount = new_total_tax_amount * 1
		If new_total_lot_rent_amount = "" Then new_total_lot_rent_amount = 0
		new_total_lot_rent_amount = new_total_lot_rent_amount * 1
		If new_total_room_amount = "" Then new_total_room_amount = 0
		new_total_room_amount = new_total_room_amount * 1
		If new_vehicle_insurance_amount = "" Then new_vehicle_insurance_amount = 0
		new_vehicle_insurance_amount = new_vehicle_insurance_amount * 1
		If new_total_insurance_amount = "" Then new_total_insurance_amount = 0
		new_total_insurance_amount = new_total_insurance_amount * 1
		If new_total_subsidy_amount = "" Then new_total_subsidy_amount = 0
		new_total_subsidy_amount = new_total_subsidy_amount * 1


		new_total_insurance_amount = new_morgage_insurance_amount + new_renter_insurance_amount + new_vehicle_insurance_amount

		new_total_shelter_expense_amount = 0
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_rent_amount
		' new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_renter_insurance_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_garage_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_mortgage_amount
		' new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_morgage_insurance_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_tax_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_lot_rent_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_room_amount
		' new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_vehicle_insurance_amount
		new_total_shelter_expense_amount = new_total_shelter_expense_amount + new_total_insurance_amount

		new_total_shelter_expense_amount = new_total_shelter_expense_amount - new_total_subsidy_amount

		number_of_people_paying = 0
		for the_membs = 0 to UBound(THE_ARRAY, 2)
			If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then number_of_people_paying = number_of_people_paying + 1
		Next
		If other_person_checkbox = checked Then number_of_people_paying = number_of_people_paying + 1
		If number_of_people_paying = 1 Then
			for the_membs = 0 to UBound(THE_ARRAY, 2)
				If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
					THE_ARRAY(new_shel_pers_total_amt_const, the_membs) = new_total_shelter_expense_amount & ""
					THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs) = "dollars"
				End If
			Next
			If other_person_checkbox = checked Then
				other_new_shel_total_amt = new_total_shelter_expense_amount & ""
				other_new_shel_total_amt_type = "dollars"
			End If
		ElseIf payment_split_evenly_yn = "Yes" Then
			If number_of_people_paying <> 0 Then
				amount_per_person = new_total_shelter_expense_amount/number_of_people_paying
				for the_membs = 0 to UBound(THE_ARRAY, 2)
					If THE_ARRAY(person_shel_checkbox, the_membs) = checked Then
						THE_ARRAY(new_shel_pers_total_amt_const, the_membs) = amount_per_person & ""
						THE_ARRAY(new_shel_pers_total_amt_type_const, the_membs) = "dollars"
					End If
				Next
				If other_person_checkbox = checked Then
					other_new_shel_total_amt = amount_per_person & ""
					other_new_shel_total_amt_type = "dollars"
				End If
			End If
		End if

		new_total_rent_amount = new_total_rent_amount & ""
		new_renter_insurance_amount = new_renter_insurance_amount & ""
		new_total_garage_amount = new_total_garage_amount & ""
		new_total_mortgage_amount = new_total_mortgage_amount & ""
		new_morgage_insurance_amount = new_morgage_insurance_amount & ""
		new_total_tax_amount = new_total_tax_amount & ""
		new_total_lot_rent_amount = new_total_lot_rent_amount & ""
		new_total_room_amount = new_total_room_amount & ""
		new_vehicle_insurance_amount = new_vehicle_insurance_amount & ""
		new_total_insurance_amount = new_total_insurance_amount & ""

	End If

	view_shel_details_dlg = False
	If household_move_yn = "?" Then
		view_addr_update_dlg = "Unknown"
		err_msg = "STOP"
	End If

	If shel_change_yn = "?" Then
		view_shel_update_dlg = "Unknown"
		err_msg = "STOP"
	End If

	If household_move_yn = "Yes" Then
		view_addr_update_dlg = True
		shel_change_yn = "Yes"
		view_shel_update_dlg = True
		If shel_shared_yn = "Yes" Then view_shel_details_dlg = True
	End If
	If household_move_yn = "No" Then
		view_addr_update_dlg = False
		view_shel_details_dlg = False
	End If

	' household_move_yn,
	' household_move_everyone_yn,
	' move_date,
	' shel_change_yn,
	' shel_verif_received_yn,
	' shel_start_date,
	' shel_shared_yn,
	' shel_subsidized_yn,
	' total_current_rent,
	' total_current_taxes,
	' total_current_lot_rent,
	' total_current_room,
	' total_current_mortgage,
	' total_current_garage,
	' total_current_insurance,
	' total_current_subsidy,
	' shel_change_type,
	' hest_heat_ac_yn,
	' hest_electric_yn,
	' hest_ac_on_electric_yn,
	' hest_heat_on_electric_yn,
	' hest_phone_yn


	' MsgBox err_msg
	If err_msg = "" Then
		If ButtonPressed = housing_change_continue_btn Then
			housing_questions_step = housing_questions_step + 1

			If housing_questions_step = 2 and view_addr_update_dlg = False Then housing_questions_step = housing_questions_step + 1
			If housing_questions_step = 3 and view_shel_update_dlg = False Then housing_questions_step = housing_questions_step + 1
			' If housing_questions_step = 4 and view_shel_details_dlg = False Then housing_questions_step = housing_questions_step + 1

			' If housing_questions_step = 1 Then 		'Initial Basic questions
			' ElseIf housing_questions_step = 2 Then 		'update ADDR Information
			'
			' ElseIf housing_questions_step = 3 Then 		'update SHEL Information
			'
			' ElseIf housing_questions_step = 4 Then 		'update SHEL percentages and SHARING Information
			'
			' ElseIf housing_questions_step = 5 Then 		'review information
			'
			' End If
		End If
		If ButtonPressed = housing_change_overview_btn Then housing_questions_step = 1
		If ButtonPressed = housing_change_addr_update_btn Then housing_questions_step = 2
		If ButtonPressed = housing_change_shel_update_btn Then housing_questions_step = 3
		' If ButtonPressed = housing_change_shel_details_btn Then housing_questions_step = 4

		If housing_questions_step = 3 Then

			if start_on_shel_questions = False Then shel_update_step = 1

			If ButtonPressed = enter_shel_one_btn Then shel_update_step = 2
			If ButtonPressed = enter_shel_two_btn Then shel_update_step = 3
			If ButtonPressed = enter_shel_three_btn Then shel_update_step = 4
			' If ButtonPressed = enter_shel_two_btn Then shel_update_step = 2
			' If ButtonPressed = enter_shel_two_btn Then shel_update_step = 2
			' If ButtonPressed = enter_shel_two_btn Then shel_update_step = 2

			total_current_rent = total_current_rent & ""
			total_current_lot_rent = total_current_lot_rent & ""
			total_current_garage = total_current_garage & ""
			total_current_insurance = total_current_insurance & ""
			total_current_taxes = total_current_taxes & ""
			total_current_room = total_current_room & ""
			total_current_mortgage = total_current_mortgage & ""
			total_current_subsidy = total_current_subsidy & ""
			' all_rent_verif,
			' all_lot_rent_verif,
			' all_mortgage_verif,
			' all_insurance_verif,
			' all_taxes_verif,
			' all_room_verif,
			' all_garage_verif,
			' all_subsidy_verif,
			' total_shel_original_information)

		End If
	End If

	' MsgBox housing_questions_step
end function

'==========================================================================================================================

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

BeginDialog Dialog1, 0, 0, 196, 60, "Dialog"
  DropListBox 15, 15, 170, 45, " "+chr(9)+"Application/Renewal"+chr(9)+"Change", select_option
  ButtonGroup ButtonPressed
    OkButton 135, 35, 50, 15
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

		If page_to_display = HEST_dlg_page Then Call navigate_HEST_buttons(update_hest, err_msg, update_information_btn, save_information_btn, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense, date)

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
	If hest_update_attempted = True Then Call access_HEST_panel("WRITE", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)



End If


If select_option = "Change" Then

	Call read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, total_shelter_expense, total_shel_original_information)
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_addr_panel_info, addr_update_attempted)
	Call access_HEST_panel("READ", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

	housing_questions_step = 1
	view_addr_update_dlg = False
	view_shel_update_dlg = False
	view_shel_details_dlg = False
	page_to_display = CHNG_dlg_page

	Do
		err_msg = ""

		MsgBox "HOUSING QUESTIONS - " & housing_questions_step & vbCr & "SHEL UPDATE - " & shel_update_step

		BeginDialog Dialog1, 0, 0, 555, 385, "Housing Expense Detail"

		  ButtonGroup ButtonPressed

			If page_to_display = CHNG_dlg_page Then
				Text 503, 57, 60, 10, "CHANGE"
				Call display_HOUSING_CHANGE_information(housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, ALL_SHEL_PANELS_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn)

			End If

			If page_to_display <> CHNG_dlg_page Then PushButton 485, 55, 65, 13, "CHANGE", CHNG_page_btn

			OkButton 450, 365, 50, 15
			CancelButton 500, 365, 50, 15

		EndDialog


		Dialog Dialog1
		cancel_without_confirmation
		' MsgBox "Button - " & ButtonPressed

		If ButtonPressed = -1 AND housing_questions_step <> 5 Then ButtonPressed = housing_change_continue_btn
		If page_to_display = CHNG_dlg_page Then Call navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, shel_update_step, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, what_is_the_living_arrangement, unit_owned, new_total_rent_amount, new_total_mortgage_amount, new_total_lot_rent_amount, new_total_room_amount, new_room_payment_frequency, new_mortgage_have_escrow_yn, new_morgage_insurance_amount, new_excess_insurance_yn, new_total_tax_amount, new_rent_subsidy_yn, new_renter_insurance_amount, new_renters_insurance_required_yn, new_total_garage_amount, new_garage_rent_required_yn, new_vehicle_insurance_amount, new_total_insurance_amount, new_total_subsidy_amount, new_SHEL_paid_to_name, other_person_checkbox, other_person_name, payment_split_evenly_yn, ALL_SHEL_PANELS_ARRAY, person_age_const, person_shel_checkbox, shel_ref_number_const, new_shel_pers_total_amt_const, new_shel_pers_total_amt_type_const, other_new_shel_total_amt, other_new_shel_total_amt_type, new_total_shelter_expense_amount, new_rent_verif, new_lot_rent_verif, new_mortgage_verif, new_insurance_verif, new_taxes_verif, new_room_verif, new_garage_verif, new_subsidy_verif, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn, enter_shel_one_btn, enter_shel_two_btn, enter_shel_three_btn, update_shel_button, addr_update_needed, shel_update_needed, hest_update_needed)
		If ButtonPressed = CHNG_page_btn Then page_to_display = CHNG_dlg_page
		If err_msg <> "" then MsgBox "Please Resolve:" & vbCr & err_msg
	Loop until ButtonPressed = -1




End If


script_end_procedure("Done")
