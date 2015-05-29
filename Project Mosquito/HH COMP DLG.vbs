DIM num_of_peeps
DIM hh_array

FUNCTION hh_comp_dlg(hh_array)
	dlg_height = 95
	start_row = 55
	num_of_peeps = 1
	ReDim hh_array(num_of_peeps - 1, 5)
	
	DO
		BeginDialog hh_comp_dlg, 0, 0, 441, dlg_height, "Household Composition"
	    	Text 10, 15, 110, 10, "Number of people in household..."
	    	EditBox 130, 10, 30, 15, num_of_peeps
		Text 10, 35, 60, 10, "Individual Name"
		Text 90, 35, 45, 10, "Date of Birth"
		Text 165, 35, 50, 10, "Relationship"
		Text 245, 35, 50, 10, "P/P Together?"
		Text 305, 35, 70, 10, "Additional Info"
    		EditBox 10, 50, 70, 15, hh_array(hh_size, 0)								'Name
    		EditBox 90, 50, 60, 15, hh_array(hh_size, 1)								'Age
    		DropListBox 160, 50, 70, 15, "Applicant", hh_array(hh_size, 2)					'Relationship to Applicant
    		DropListBox 245, 50, 55, 15, "Yes", hh_array(hh_size, 3)						'Purchase & Prepare w/ Applicant?
		CheckBox 310, 50, 55, 15, "Has Income?", hh_array(hh_size, 4)					'Does person have income?
		CheckBox 375, 50, 55, 15, "Has Assets?", hh_array(hh_size, 5)					'Does person have assets?
		IF num_of_peeps <> 1 THEN 
		    	FOR hh_size = 1 to (num_of_peeps - 1)
	    			rel_list = "Select one..."+chr(9)+"Spouse"+chr(9)+"Child"+chr(9)+"Parent"
    				purch_prepare_list = "Select one..."+chr(9)+"Yes"+chr(9)+"No"

		    		EditBox 10, start_row + 15, 70, 15, hh_array(hh_size, 0)								'Name
		    		EditBox 90, start_row + 15, 60, 15, hh_array(hh_size, 1)								'Age
	    			DropListBox 160, start_row + 15, 70, 15, rel_list, hh_array(hh_size, 2)						'Relationship to Applicant
	    			DropListBox 245, start_row + 15, 55, 15, purch_prepare_list, hh_array(hh_size, 3)				'Purchase & Prepare w/ Applicant?
				CheckBox 310, start_row + 15, 55, 15, "Has Income?", hh_array(hh_size, 4)					'Does person have income?
				CheckBox 375, start_row + 15, 55, 15, "Has Assets?", hh_array(hh_size, 5)					'Does person have assets?
		    		start_row = start_row + 15
		    	NEXT
		END IF
	    	ButtonGroup ButtonPressed
	    		PushButton 170, 10, 100, 15, "Update for household size...", UpdateButton
	    		OkButton 295, start_row + 20, 55, 15
	    		CancelButton 350, start_row + 20, 55, 15
		EndDialog

		Dialog hh_comp_dlg
			IF ButtonPressed = 0 THEN stopscript
			IF ButtonPressed = UpdateButton THEN 
				ReDim hh_array(num_of_peeps - 1, 5)
				dlg_height = 75 + (20 * num_of_peeps)
				start_row = 50
			END IF

	LOOP UNTIL ButtonPressed = -1
END FUNCTION

CALL hh_comp_dlg(hh_array)

