function diet_dialog()
	EditBox 395, 0, 45, 15, diet_date_received
	DropListBox 50, 35, 120, 15, HH_Memb_DropDown, diet_member_number
	DropListBox 55, 70, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_1_dropdown
	DropListBox 185, 70, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_1_dropdown
	DropListBox 290, 70, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_1_dropdown
	DropListBox 55, 85, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_2_dropdown
	DropListBox 185, 85, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_2_dropdown
 	DropListBox 290, 85, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_2_dropdown
	DropListBox 55, 100, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_3_dropdown
	DropListBox 185, 100, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_3_dropdown
	DropListBox 290, 100, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_3_dropdown
	DropListBox 55, 115, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_4_dropdown
	DropListBox 185, 115, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_4_dropdown
	DropListBox 290, 115, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_4_dropdown
	DropListBox 55, 130, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_5_dropdown
	DropListBox 185, 130, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_5_dropdown
	DropListBox 290, 130, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_5_dropdown
	DropListBox 55, 145, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_6_dropdown
	DropListBox 185, 145, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_6_dropdown
	DropListBox 290, 145, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_6_dropdown
	DropListBox 55, 160, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_7_dropdown
	DropListBox 185, 160, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_7_dropdown
	DropListBox 290, 160, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_7_dropdown
	DropListBox 55, 175, 115, 20, ""+chr(9)+"01-High Protein"+chr(9)+"02-Controlled protein 40-60 grams"+chr(9)+"03-Controlled protein <40 grams"+chr(9)+"04-Low cholesterol"+chr(9)+"05-High residue"+chr(9)+"06-Pregnancy/Lactation"+chr(9)+"07-Gluten free"+chr(9)+"08-Lactose free"+chr(9)+"09-Anti-dumping"+chr(9)+"10-Hypoglycemic"+chr(9)+"11-Ketogenic", diet_8_dropdown
	DropListBox 185, 175, 90, 15, ""+chr(9)+"N/A - Only 1 diet"+chr(9)+"Non-Overlapping"+chr(9)+"Overlapping"+chr(9)+"Mutually Exclusive"+chr(9)+"Blank", diet_relationship_8_dropdown
	DropListBox 290, 175, 30, 15, ""+chr(9)+"Yes"+chr(9)+"No", diet_verif_8_dropdown
	EditBox 75, 195, 55, 15, diet_date_last_exam
	DropListBox 135, 215, 35, 15, ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank", diet_treatment_plan_dropdown
	EditBox 270, 215, 55, 15, diet_length_diet
	DropListBox 55, 235, 60, 15, ""+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Incomplete", diet_status_dropdown
	CheckBox 125, 235, 195, 10, "Check here to set TIKL for renewal", diet_tikl_checkbox
	EditBox 50, 260, 290, 15, diet_comments
	PushButton 5, 280, 80, 15, "CM23.12- Special Diets", diet_link_CM_special_diet
    PushButton 95, 280, 115, 15, "Processing Special Diet Referrals", diet_SP_referrals
	Text 5, 5, 220, 10, diet_form_name
	Text 340, 5, 55, 10, "Document Date:"
	Text 20, 40, 30, 10, "Member"
	Text 20, 20, 325, 10, diet_mfip_msa_status
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

Do							'Do Loop to cycle through dialog as many times as needed until all desired forms are added
	Do
		Do
			err_msg = ""
			Dialog1 = "" 			'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 296, 290, "Select Documents Received"
				DropListBox 30, 30, 180, 15, ""+chr(9)+asset_form_name+chr(9)+atr_form_name+chr(9)+arep_form_name+chr(9)+change_form_name+chr(9)+evf_form_name+chr(9)+hosp_form_name+chr(9)+iaa_form_name+chr(9)+ltc_1503_form_name+chr(9)+mof_form_name+chr(9)+psn_form_name+chr(9)+sf_form_name+chr(9)+diet_form_name+chr(9)+other_form_name, Form_type
				ButtonGroup ButtonPressed
				PushButton 225, 30, 35, 10, "Add", add_button
				PushButton 225, 60, 35, 10, "All Forms", all_forms
				PushButton 155, 270, 40, 15, "Clear", clear_button
				OkButton 205, 270, 40, 15
				CancelButton 255, 270, 40, 15
				GroupBox 5, 5, 280, 70, "Directions: For each document received either:"
				Text 15, 15, 275, 10, "1. Select document from dropdown, then select Add button. Repeat for each form."
				Text 10, 45, 15, 10, "OR"
				Text 15, 60, 180, 10, "2. Select All Forms to select forms via checkboxes."
				GroupBox 45, 85, 210, 175, "Documents Selected"
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
	form_type = ""	'Resets the form value to blank after each selection

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
					CheckBox 15, 20, 170, 10, asset_form_name, asset_checkbox
					CheckBox 15, 30, 170, 10, atr_form_name, atr_checkbox
					CheckBox 15, 40, 170, 10, arep_form_name, arep_checkbox
					CheckBox 15, 50, 170, 10, change_form_name, change_checkbox
					CheckBox 15, 60, 170, 10, evf_form_name, evf_checkbox
					CheckBox 15, 70, 170, 10, hosp_form_name, hospice_checkbox
					CheckBox 15, 80, 170, 10, iaa_form_name, iaa_checkbox
					CheckBox 15, 90, 170, 10, ltc_1503_form_name, ltc_1503_checkbox
					CheckBox 15, 100, 170, 10, mof_form_name, mof_checkbox
					CheckBox 15, 110, 170, 10, psn_form_name, psn_checkbox
					CheckBox 15, 120, 170, 10, sf_form_name, shelter_checkbox
					CheckBox 15, 130, 170, 10, diet_form_name, diet_checkbox
					CheckBox 15, 140, 170, 10, other_form_name, other_checkbox
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
				If asset_checkbox = unchecked and arep_checkbox = unchecked and atr_checkbox = unchecked and change_checkbox = unchecked and evf_checkbox = unchecked and hospice_checkbox = unchecked and iaa_checkbox = unchecked and ltc_1503_checkbox = unchecked and mof_checkbox = unchecked and psn_checkbox = unchecked and shelter_checkbox = unchecked and diet_checkbox = unchecked and other_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "-Select forms to process or select cancel to exit script"		'If review selections is selected and all checkboxes are blank, user will receive error
				If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg							'list of errors to resolve
			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

	End If
Loop Until ButtonPressed = Ok