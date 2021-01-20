'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - COMPLETE PHONE CSR.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================
' run_locally = TRUE
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

'FUNCTIONS =================================================================================================================

function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")

        EMReadScreen line_one, 22, 6, 43
        EMReadScreen line_two, 22, 7, 43
        EMReadScreen city_line, 15, 8, 43
        EMReadScreen state_line, 2, 8, 66
        EMReadScreen zip_line, 7, 9, 43
        EMReadScreen county_line, 2, 9, 66
        EMReadScreen verif_line, 2, 9, 74
        EMReadScreen homeless_line, 1, 10, 43
        EMReadScreen reservation_line, 1, 10, 74
        EMReadScreen living_sit_line, 2, 11, 43

        resi_line_one = replace(line_one, "_", "")
        resi_line_two = replace(line_two, "_", "")
        resi_city = replace(city_line, "_", "")
        resi_state = state_line
        resi_zip = replace(zip_line, "_", "")

        If county_line = "01" Then addr_county = "01 Aitkin"
        If county_line = "02" Then addr_county = "02 Anoka"
        If county_line = "03" Then addr_county = "03 Becker"
        If county_line = "04" Then addr_county = "04 Beltrami"
        If county_line = "05" Then addr_county = "05 Benton"
        If county_line = "06" Then addr_county = "06 Big Stone"
        If county_line = "07" Then addr_county = "07 Blue Earth"
        If county_line = "08" Then addr_county = "08 Brown"
        If county_line = "09" Then addr_county = "09 Carlton"
        If county_line = "10" Then addr_county = "10 Carver"
        If county_line = "11" Then addr_county = "11 Cass"
        If county_line = "12" Then addr_county = "12 Chippewa"
        If county_line = "13" Then addr_county = "13 Chisago"
        If county_line = "14" Then addr_county = "14 Clay"
        If county_line = "15" Then addr_county = "15 Clearwater"
        If county_line = "16" Then addr_county = "16 Cook"
        If county_line = "17" Then addr_county = "17 Cottonwood"
        If county_line = "18" Then addr_county = "18 Crow Wing"
        If county_line = "19" Then addr_county = "19 Dakota"
        If county_line = "20" Then addr_county = "20 Dodge"
        If county_line = "21" Then addr_county = "21 Douglas"
        If county_line = "22" Then addr_county = "22 Faribault"
        If county_line = "23" Then addr_county = "23 Fillmore"
        If county_line = "24" Then addr_county = "24 Freeborn"
        If county_line = "25" Then addr_county = "25 Goodhue"
        If county_line = "26" Then addr_county = "26 Grant"
        If county_line = "27" Then addr_county = "27 Hennepin"
        If county_line = "28" Then addr_county = "28 Houston"
        If county_line = "29" Then addr_county = "29 Hubbard"
        If county_line = "30" Then addr_county = "30 Isanti"
        If county_line = "31" Then addr_county = "31 Itasca"
        If county_line = "32" Then addr_county = "32 Jackson"
        If county_line = "33" Then addr_county = "33 Kanabec"
        If county_line = "34" Then addr_county = "34 Kandiyohi"
        If county_line = "35" Then addr_county = "35 Kittson"
        If county_line = "36" Then addr_county = "36 Koochiching"
        If county_line = "37" Then addr_county = "37 Lac Qui Parle"
        If county_line = "38" Then addr_county = "38 Lake"
        If county_line = "39" Then addr_county = "39 Lake Of Woods"
        If county_line = "40" Then addr_county = "40 Le Sueur"
        If county_line = "41" Then addr_county = "41 Lincoln"
        If county_line = "42" Then addr_county = "42 Lyon"
        If county_line = "43" Then addr_county = "43 Mcleod"
        If county_line = "44" Then addr_county = "44 Mahnomen"
        If county_line = "45" Then addr_county = "45 Marshall"
        If county_line = "46" Then addr_county = "46 Martin"
        If county_line = "47" Then addr_county = "47 Meeker"
        If county_line = "48" Then addr_county = "48 Mille Lacs"
        If county_line = "49" Then addr_county = "49 Morrison"
        If county_line = "50" Then addr_county = "50 Mower"
        If county_line = "51" Then addr_county = "51 Murray"
        If county_line = "52" Then addr_county = "52 Nicollet"
        If county_line = "53" Then addr_county = "53 Nobles"
        If county_line = "54" Then addr_county = "54 Norman"
        If county_line = "55" Then addr_county = "55 Olmsted"
        If county_line = "56" Then addr_county = "56 Otter Tail"
        If county_line = "57" Then addr_county = "57 Pennington"
        If county_line = "58" Then addr_county = "58 Pine"
        If county_line = "59" Then addr_county = "59 Pipestone"
        If county_line = "60" Then addr_county = "60 Polk"
        If county_line = "61" Then addr_county = "61 Pope"
        If county_line = "62" Then addr_county = "62 Ramsey"
        If county_line = "63" Then addr_county = "63 Red Lake"
        If county_line = "64" Then addr_county = "64 Redwood"
        If county_line = "65" Then addr_county = "65 Renville"
        If county_line = "66" Then addr_county = "66 Rice"
        If county_line = "67" Then addr_county = "67 Rock"
        If county_line = "68" Then addr_county = "68 Roseau"
        If county_line = "69" Then addr_county = "69 St. Louis"
        If county_line = "70" Then addr_county = "70 Scott"
        If county_line = "71" Then addr_county = "71 Sherburne"
        If county_line = "72" Then addr_county = "72 Sibley"
        If county_line = "73" Then addr_county = "73 Stearns"
        If county_line = "74" Then addr_county = "74 Steele"
        If county_line = "75" Then addr_county = "75 Stevens"
        If county_line = "76" Then addr_county = "76 Swift"
        If county_line = "77" Then addr_county = "77 Todd"
        If county_line = "78" Then addr_county = "78 Traverse"
        If county_line = "79" Then addr_county = "79 Wabasha"
        If county_line = "80" Then addr_county = "80 Wadena"
        If county_line = "81" Then addr_county = "81 Waseca"
        If county_line = "82" Then addr_county = "82 Washington"
        If county_line = "83" Then addr_county = "83 Watonwan"
        If county_line = "84" Then addr_county = "84 Wilkin"
        If county_line = "85" Then addr_county = "85 Winona"
        If county_line = "86" Then addr_county = "86 Wright"
        If county_line = "87" Then addr_county = "87 Yellow Medicine"
        If county_line = "89" Then addr_county = "89 Out-of-State"
        resi_county = addr_county

        If homeless_line = "Y" Then addr_homeless = "Yes"
        If homeless_line = "N" Then addr_homeless = "No"
        If reservation_line = "Y" Then addr_reservation = "Yes"
        If reservation_line = "N" Then addr_reservation = "No"

        If verif_line = "SF" Then addr_verif = "SF - Shelter Form"
        If verif_line = "Co" Then addr_verif = "CO - Coltrl Stmt"
        If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
        If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
        If verif_line = "CD" Then addr_verif = "CD - Contrct for Deed"
        If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
        If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
        If verif_line = "OT" Then addr_verif = "OT - Other Document"
        If verif_line = "NO" Then addr_verif = "NO - No Ver Prvd"
        If verif_line = "?_" Then addr_verif = "? - Delayed"
        If verif_line = "__" Then addr_verif = "Blank"


        If living_sit_line = "__" Then living_situation = "Blank"
        If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roomate"
        If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
        If living_sit_line = "03" Then living_situation = "03 -  servc prvdr- foster/group home"
        If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
        If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
        If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
        If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
        If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
        If living_sit_line = "09" Then living_situation = "09 - Declined"
        If living_sit_line = "10" Then living_situation = "10 - Unknown"
        addr_living_sit = living_situation

        EMReadScreen addr_eff_date, 8, 4, 43
        EMReadScreen addr_future_date, 8, 4, 66
        EMReadScreen mail_line_one, 22, 13, 43
        EMReadScreen mail_line_two, 22, 14, 43
        EMReadScreen mail_city_line, 15, 15, 43
        EMReadScreen mail_state_line, 2, 16, 43
        EMReadScreen mail_zip_line, 7, 16, 52

        addr_eff_date = replace(addr_eff_date, " ", "/")
        addr_future_date = trim(addr_future_date)
        addr_future_date = replace(addr_future_date, " ", "/")
        mail_line_one = replace(mail_line_one, "_", "")
        mail_line_two = replace(mail_line_two, "_", "")
        mail_city = replace(mail_city_line, "_", "")
        mail_state = replace(mail_state_line, "_", "")
        mail_zip = replace(mail_zip_line, "_", "")

        notes_on_address = "Address effective: " & addr_eff_date & "."
        ' If mail_line_one <> "" Then
        '     If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        '     If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
        ' End If
        If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."

        EMReadScreen phone_one, 14, 17, 45
        EMReadScreen phone_two, 14, 18, 45
        EMReadScreen phone_three, 14, 19, 45

        EMReadScreen type_one, 1, 17, 67
        EMReadScreen type_two, 1, 18, 67
        EMReadScreen type_three, 1, 19, 67

        phone_one = replace(phone_one, " ) ", "-")
        phone_one = replace(phone_one, " ", "-")
        If phone_one = "___-___-____" Then phone_one = ""

        phone_two = replace(phone_two, " ) ", "-")
        phone_two = replace(phone_two, " ", "-")
        If phone_two = "___-___-____" Then phone_two = ""

        phone_three = replace(phone_three, " ) ", "-")
        phone_three = replace(phone_three, " ", "-")
        If phone_three = "___-___-____" Then phone_three = ""

        If type_one = "H" Then type_one = "H - Home"
        If type_one = "W" Then type_one = "W - Work"
        If type_one = "C" Then type_one = "C - Cell"
        If type_one = "M" Then type_one = "M - Message"
        If type_one = "T" Then type_one = "T - TTY/TDD"
        If type_one = "_" Then type_one = ""

        If type_two = "H" Then type_two = "H - Home"
        If type_two = "W" Then type_two = "W - Work"
        If type_two = "C" Then type_two = "C - Cell"
        If type_two = "M" Then type_two = "M - Message"
        If type_two = "T" Then type_two = "T - TTY/TDD"
        If type_two = "_" Then type_two = ""

        If type_three = "H" Then type_three = "H - Home"
        If type_three = "W" Then type_three = "W - Work"
        If type_three = "C" Then type_three = "C - Cell"
        If type_three = "M" Then type_three = "M - Message"
        If type_three = "T" Then type_three = "T - TTY/TDD"
        If type_three = "_" Then type_three = ""
    End If
end function

function split_phone_number_into_parts(phone_variable, phone_left, phone_mid, phone_right)
    phone_variable = trim(phone_variable)
    If phone_variable <> "" Then
        phone_variable = replace(phone_variable, "(", "")
        phone_variable = replace(phone_variable, ")", "")
        phone_variable = replace(phone_variable, "-", "")
        phone_variable = replace(phone_variable, " ", "")
        phone_variable = trim(phone_variable)
        phone_left = left(phone_variable, 3)
        phone_mid = mid(phone_variable, 4, 3)
        phone_right = right(phone_variable, 4)
        phone_variable = "(" & phone_left & ")" & phone_mid & "-" & phone_right
    End If
end function

'THESE FUNCTIONS CREATE, DECLARE, AND SHOW MOST OF THE DIALOGS
function csr_dlg_q_1()
	Do
		Do
			'This dialog reviews address and household composition
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 361, 245, "CSR Detail and Address"
			  GroupBox 5, 5, 350, 30, "CSR Name Information"
			  Text 15, 20, 130, 10, "Who are you completing the form with?"
			  ComboBox 150, 15, 150, 45, all_the_clients+chr(9)+"Person Information Missing", client_on_csr_form
			  GroupBox 5, 40, 350, 60, "SR Programs"
			  Text 15, 55, 155, 10, "Are you processing a SNAP six-month report?"
			  DropListBox 165, 50, 40, 45, "Yes"+chr(9)+"No", snap_sr_yn
			  Text 220, 55, 40, 10, "Month/Year"
			  EditBox 260, 50, 20, 15, snap_sr_mo
			  EditBox 285, 50, 20, 15, snap_sr_yr
			  Text 15, 70, 155, 10, "Are you processing a HC six-month report?"
			  DropListBox 165, 65, 40, 45, "Yes"+chr(9)+"No", hc_sr_yn
			  Text 220, 70, 40, 10, "Month/Year"
			  EditBox 260, 65, 20, 15, hc_sr_mo
			  EditBox 285, 65, 20, 15, hc_sr_yr
			  Text 15, 85, 155, 10, "Are you processing a GRH six-month report?"
			  DropListBox 165, 80, 40, 45, "Yes"+chr(9)+"No", grh_sr_yn
			  Text 220, 85, 40, 10, "Month/Year"
			  EditBox 260, 80, 20, 15, grh_sr_mo
			  EditBox 285, 80, 20, 15, grh_sr_yr
			  GroupBox 5, 105, 350, 115, "Address on CSR"
			  Text 10, 120, 310, 10, "Ask client for their address. Compare it to the current addresses listed below from MAXIS:"
			  If new_resi_addr_entered = TRUE Then
				  Text 10, 135, 110, 10, "UPDATED Residence Address:"
				  Text 20, 145, 110, 10, new_resi_one
				  Text 20, 155, 115, 10, new_resi_city & ", " & new_resi_state & " " & new_resi_zip
			  Else
				  Text 10, 135, 75, 10, "Residence Address:"
				  Text 20, 145, 110, 10, resi_line_one
				  If resi_line_two = "" Then
					  Text 20, 155, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
				  Else
					  Text 20, 155, 115, 10, resi_line_two
					  Text 20, 165, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
				  End If
			  End If
			  If new_mail_addr_entered = TRUE Then
				  Text 205, 135, 110, 10, "UPDATED Mailing Address:"
				  Text 210, 145, 110, 10, new_mail_one
				  Text 210, 155, 120, 10, new_mail_city & ", " & new_mail_state & " " & new_mail_zip
			  Else
		 		  Text 205, 135, 75, 10, "Mailing Address:"
				  If mail_line_one = "" Then
					  Text 210, 145, 110, 10, "NO MAILING ADDRESS LISTED"
				  Else
					  Text 210, 145, 110, 10, mail_line_one
					  If mail_line_two = "" Then
						  Text 210, 155, 120, 10, mail_city & ", " & mail_state & " " & mail_zip
					  Else
						  Text 210, 155, 120, 10, mail_line_two
						  Text 210, 165, 120, 10, mail_city & ", " & mail_state & " " & mail_zip
					  End If
				  End If
			  End IF
			  Text 20, 205, 145, 10, "Is the Client indicating they are homeless?"
			  DropListBox 165, 200, 85, 45, "Select One..."+chr(9)+"Yes - Homeless"+chr(9)+"No", homeless_status
			  ButtonGroup ButtonPressed
			    PushButton 10, 185, 120, 10, "RESI ADDRESS IS DIFFERENT", change_resi_addr_btn
			    PushButton 205, 185, 120, 10, "MAIL ADDRESS IS DIFFERENT", change_mail_addr_btn
			    OkButton 250, 225, 50, 15
			    CancelButton 305, 225, 50, 15
			EndDialog

			err_msg = "LOOP"

		    dialog Dialog1
		    cancel_confirmation

			If ButtonPressed = change_resi_addr_btn Then call enter_new_residence_address
			If ButtonPressed = change_mail_addr_btn Then call enter_new_mailing_address

		    program_indicated = FALSE
		    If snap_sr_yn = "Yes" Then
		        Call validate_footer_month_entry(snap_sr_mo, snap_sr_yr, err_msg, "* SNAP SR MONTH")
		        program_indicated = TRUE
		    End If
		    If hc_sr_yn = "Yes" Then
		        Call validate_footer_month_entry(hc_sr_mo, hc_sr_yr, err_msg, "* HC SR MONTH")
		        program_indicated = TRUE
		    End If
		    If grh_sr_yn = "Yes" Then
		        Call validate_footer_month_entry(grh_sr_mo, grh_sr_yr, err_msg, "* GRH SR MONTH")
		        program_indicated = TRUE
		    End If

			If ButtonPressed = -1 Then
				err_msg = ""
			    If client_on_csr_form = "Select or Type" OR trim(client_on_csr_form) = "" Then err_msg = err_msg & vbNewLine & "* Indicate who is listed on the CSR form in the person infromation, or if this is blank, select that the person information is missing."
			    If program_indicated = FALSE Then err_msg = err_msg & vbNewLine & "* Select the program(s) that the CSR form is processing. (None of the programs are indicated to have an SR due.)"
			    If homeless_yn = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the CSR form indicates the household is homeless or not."

			    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
			End If
		Loop until err_msg = ""

		show_csr_dlg_q_1 = FALSE
		csr_dlg_q_1_cleared = TRUE
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
end function

function csr_dlg_q_2()
    Do

		dlg_width = 425
		If grh_sr = TRUE Then dlg_width = dlg_width + 50
		If hc_sr = TRUE Then dlg_width = dlg_width + 50
		If snap_sr = TRUE Then dlg_width = dlg_width + 50

		dlg_len = 100
		For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		    dlg_len = dlg_len + 15
		Next

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, dlg_width, dlg_len, "CSR Household"
		  Text 15, 10, 220, 10, "Q2. Has anyone moved in or out of your home in the past six months?"
		  DropListBox 240, 5, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_two_move_in_out
		  x_pos = 430
		  y_pos = 25
		  Text 15, y_pos, 300, 10, "Review the members known in MAXIS. Confirm they are still in the HH, have moved out, or recently moved in:"
		  y_pos = y_pos + 15
		  Text 15, y_pos, 35, 10, "Member #"
		  Text 60, y_pos, 40, 10, "Last Name"
		  Text 130, y_pos, 40, 10, "First Name"
		  Text 205, y_pos, 15, 10, "Age"
		  Text 295, y_pos, 50, 10, "HH Moved Out"
		  Text 360, y_pos, 55, 10, "HH Moved In"
		  If grh_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "GRH"
		      grh_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  If hc_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "HC"
		      hc_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  If snap_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "SNAP"
		      snap_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  y_pos = y_pos + 20
		  For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		    Text 20, y_pos, 15, 10, ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb)
		    Text 60, y_pos, 65, 10, ALL_CLIENTS_ARRAY(memb_last_name, known_memb)
		    Text 130, y_pos, 65, 10, ALL_CLIENTS_ARRAY(memb_first_name, known_memb)
		    Text 205, y_pos, 30, 10, ALL_CLIENTS_ARRAY(memb_age, known_memb)
		    CheckBox 305, y_pos, 25, 10, "Out", ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb)
		    CheckBox 370, y_pos, 25, 10, "In", ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb)
		    x_pos = 430
		    If grh_sr = TRUE Then Text grh_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_grh_status, known_memb)
		    If hc_sr = TRUE Then Text hc_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_hc_status, known_memb)
		    If snap_sr = TRUE Then Text snap_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_snap_status, known_memb)
		    y_pos = y_pos + 15
		  Next
		  Text 15, y_pos + 5, 275, 10, "Are there new household members that have been reported that are not listed here?"
		  DropListBox 295, y_pos, 150, 45, "Select One..."+chr(9)+"Yes - add another member"+chr(9)+"No - all member in MAXIS"+chr(9)+"New Members Have been Added", new_hh_memb_not_in_mx_yn
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
		    OkButton dlg_width - 110, y_pos, 50, 15
		    CancelButton dlg_width - 55, y_pos, 50, 15
		EndDialog

        err_msg = ""

        dialog Dialog1
        cancel_confirmation

        If quest_two_move_in_out = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the answer for Question 2 as provided on the CSR Form."
        If new_hh_memb_not_in_mx_yn = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are new members of the household that are not listed on this dialog."
		If new_hh_memb_not_in_mx_yn = "New Members Have been Added" AND new_memb_counter = 0 Then err_msg = err_msg & vbNewLine & "* No new members have been added during this script run. Select either 'Yes' or 'No'."
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
	show_csr_dlg_q_2 = FALSE
	csr_dlg_q_2_cleared = TRUE

	If new_hh_memb_not_in_mx_yn = "Yes - add another member" Then
	    Do
	        ReDim Preserve NEW_MEMBERS_ARRAY(new_memb_notes, new_memb_counter)

	        BeginDialog Dialog1, 0, 0, 255, 210, "New HH Member"
	          EditBox 55, 35, 120, 15, NEW_MEMBERS_ARRAY(new_first_name, new_memb_counter)
	          EditBox 235, 35, 15, 15, NEW_MEMBERS_ARRAY(new_mid_initial, new_memb_counter)
	          EditBox 55, 55, 120, 15, NEW_MEMBERS_ARRAY(new_last_name, new_memb_counter)
	          EditBox 210, 55, 40, 15, NEW_MEMBERS_ARRAY(new_suffix, new_memb_counter)
	          EditBox 55, 75, 50, 15, NEW_MEMBERS_ARRAY(new_dob, new_memb_counter)
	          DropListBox 105, 95, 145, 45, "Select One..."+chr(9)+"01 - Applicant"+chr(9)+"02 - Spouse"+chr(9)+"03 - Child"+chr(9)+"04 - Parent"+chr(9)+"05 - Sibling"+chr(9)+"06 - Step Sibling"+chr(9)+"08 - Step Child"+chr(9)+"09 - Step Parent"+chr(9)+"10 - Aunt"+chr(9)+"11 - Uncle"+chr(9)+"12 - Niece"+chr(9)+"13 - Nephew"+chr(9)+"14 - Cousin"+chr(9)+"15 - Grandparent"+chr(9)+"16 - Grandchild"+chr(9)+"17 - Other Relative"+chr(9)+"18 - Legal Guardian"+chr(9)+"24 - Not Related"+chr(9)+"25 - Live-In Attendant"+chr(9)+"27 - Unknown", NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter)
	          CheckBox 35, 130, 20, 10, "HC", NEW_MEMBERS_ARRAY(new_ma_request, new_memb_counter)
	          CheckBox 65, 130, 30, 10, "SNAP", NEW_MEMBERS_ARRAY(new_fs_request, new_memb_counter)
	          CheckBox 100, 130, 30, 10, "GRH", NEW_MEMBERS_ARRAY(new_grh_request, new_memb_counter)
	          CheckBox 200, 115, 50, 10, "Moved In", NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter)
	          CheckBox 200, 130, 50, 10, "Moved Out", NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter)
	          EditBox 40, 150, 210, 15, NEW_MEMBERS_ARRAY(new_memb_notes, new_memb_counter)
	          ButtonGroup ButtonPressed
	            PushButton 145, 190, 50, 15, "Add Another", add_another_new_memb_btn
	            PushButton 200, 190, 50, 15, "No More", done_adding_new_memb_btn
	          Text 10, 10, 155, 20, "Enter any information about the new household member that has not been added to MAXIS."
	          Text 10, 40, 40, 10, "First Name"
	          Text 180, 40, 45, 10, "Middle Initial:"
	          Text 10, 60, 40, 10, "Last Name:"
	          Text 180, 60, 25, 10, "Suffix:"
	          Text 10, 80, 45, 10, "Date of Birth:"
	          Text 10, 100, 85, 10, "Relationship to Memb 01:"
	          GroupBox 10, 115, 165, 30, "Check any programs this Memb is requesting"
	          Text 10, 155, 25, 10, "Notes:"
	          Text 15, 175, 95, 25, "This script will not add this information to STAT, it will CASE:NOTE the information."
	        EndDialog

	        Dialog Dialog1
	        cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = add_another_new_memb_btn
			If ButtonPressed = 0 Then ButtonPressed = done_adding_new_memb_btn
	        If ButtonPressed = add_another_new_memb_btn Then new_memb_counter = new_memb_counter + 1
	    Loop until ButtonPressed = done_adding_new_memb_btn
		new_hh_memb_not_in_mx_yn = "New Members Have been Added"

	    For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
	        NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) = trim(NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb))
	        NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) = trim(NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb))
	        If NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) = "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb)
	        If NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) <> "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) & ". " & NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb)
	        If NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) <> "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb)
	    Next
		new_memb_counter = new_memb_counter + 1
	End If
end function

function csr_dlg_q_4_7()
	Do
		dlg_len = 190
		q_4_grp_len = 15
		q_5_grp_len = 30
		q_6_grp_len = 30
		q_7_grp_len = 30
		For new_jobs_listed = 0 to UBound(NEW_EARNED_ARRAY, 2)
			If NEW_EARNED_ARRAY(earned_type, new_jobs_listed) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, new_jobs_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_5_grp_len = q_5_grp_len + 20
			End If
			If NEW_EARNED_ARRAY(earned_type, new_jobs_listed) = "JOBS"  AND NEW_EARNED_ARRAY(earned_prog_list, new_jobs_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_6_grp_len = q_6_grp_len + 20
			End If
		Next
		For new_unea_listed = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			If NEW_UNEARNED_ARRAY(unearned_type, new_unea_listed) = "UNEA"  AND NEW_UNEARNED_ARRAY(unearned_prog_list, new_unea_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_7_grp_len = q_7_grp_len + 20
			End If
		Next
		For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
			If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
				dlg_len = dlg_len + 20
				q_4_grp_len = q_4_grp_len + 20
			End If
		Next

		y_pos = 45

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, dlg_len, "MA CSR Income Questions"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 4 - 7:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 4 - 7 are not required.", all_questions_4_7_blank
		  GroupBox 15, 30, 585, q_4_grp_len, "Q4. Do you want to apply for MA for someone who is not getting coverage now?"
		  DropListBox 285, 25, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", apply_for_ma
		  ButtonGroup ButtonPressed
			PushButton 540, 30, 50, 10, "Add Another", add_memb_btn
		  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
			  If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
				  Text 35, y_pos + 5, 105, 10, "Select the Member requesting:"
				  ComboBox 145, y_pos, 195, 45, all_the_clients, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If y_pos = 45 Then y_pos = y_pos + 5

		  GroupBox 15, y_pos + 5, 585, q_5_grp_len, "Q5. Is anyone self-employed or does anyone expect to be self-employed?"
		  DropListBox 265, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_self_employed
		  y_pos = y_pos + 20

		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_busi_btn
		  first_busi= TRUE
		  For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
				  If first_busi = TRUE then
					  Text 35, y_pos, 25, 10, "Name"
					  Text 155, y_pos, 55, 10, "Business Name"
					  Text 265, y_pos, 35, 10, "Start Date"
					  Text 325, y_pos, 50, 10, "Yearly Income"
					  y_pos = y_pos + 10
					  first_busi = FALSE
				  End If

				  ComboBox 35, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, each_busi)
				  EditBox 155, y_pos, 105, 15, NEW_EARNED_ARRAY(earned_source, each_busi)
				  EditBox 265, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_start_date, each_busi)
				  EditBox 325, y_pos, 60, 15, NEW_EARNED_ARRAY(earned_amount, each_busi)
				  y_pos = y_pos  + 20
			  End If
		  Next
		  If first_busi = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no BUSI information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  GroupBox 15, y_pos + 5, 585, q_6_grp_len, "Q6. Does anyone work or does anyone expect to start working?"
		  DropListBox 230, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_start_working
		  y_pos = y_pos  + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_jobs_btn
		  first_job = TRUE
		  For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
				  If first_job = TRUE Then
					  Text 40, y_pos, 20, 10, "Name"
					  Text 160, y_pos, 55, 10, "Employer Name"
					  Text 270, y_pos, 35, 10, "Start Date"
					  Text 330, y_pos, 35, 10, "Seasonal"
					  Text 375, y_pos, 30, 10, "Amount"
					  Text 425, y_pos, 50, 10, "How often?"
					  y_pos = y_pos  + 10
					  first_job = FALSE
				  End If
				  ComboBox 40, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, each_job)
				  EditBox 155, y_pos, 105, 15, NEW_EARNED_ARRAY(earned_source, each_job)
				  EditBox 270, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_start_date, each_job)
				  DropListBox 330, y_pos, 40, 45, " "+chr(9)+"No"+chr(9)+"Yes", NEW_EARNED_ARRAY(earned_seasonal, each_job)
				  EditBox 375, y_pos, 45, 15, NEW_EARNED_ARRAY(earned_amount, each_job)
				  DropListBox 425, y_pos, 60, 45, "Select One..."+chr(9)+"4 - Weekly"+chr(9)+"3 - Biweekly"+chr(9)+"2 - Semi Monthly"+chr(9)+"1 - Monthly", NEW_EARNED_ARRAY(earned_freq, each_job)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_job = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no JOBS information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  GroupBox 15, y_pos + 5, 585, q_7_grp_len, "Q7. Does anyone get money or does anyone expect to get money from sources other than work?"
		  DropListBox 335, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_other_income
		  y_pos = y_pos +20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_unea_btn
		  first_unea = TRUE

		  For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			  If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
				  If first_unea = TRUE Then
					  Text 30, y_pos, 25, 10, "Name"
					  Text 165, y_pos, 55, 10, "Type of Income"
					  Text 280, y_pos, 35, 10, "Start Date"
					  Text 335, y_pos, 35, 10, "Amount"
					  Text 390, y_pos, 55, 10, "How often recvd"
					  y_pos = y_pos + 10
					  first_unea = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_UNEARNED_ARRAY(unearned_client, each_unea)
				  ComboBox 165, y_pos, 110, 45, "Type or Select"+chr(9)+UNEA_type_list, NEW_UNEARNED_ARRAY(unearned_source, each_unea)   'unea_type
				  EditBox 280, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)    'unea_start_date
				  EditBox 335, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_amount, each_unea)    'unea_amount
				  DropListBox 390, y_pos, 90, 45, "Select One..."+chr(9)+"4 - Weekly"+chr(9)+"3 - Biweekly"+chr(9)+"2 - Semi Monthly"+chr(9)+"1 - Monthly", NEW_UNEARNED_ARRAY(unearned_freq, each_unea) 'unea_frequency
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_unea = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no UNEA information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  ButtonGroup ButtonPressed
			PushButton 475, y_pos, 80, 15, "Go to Q9 - Q12", next_page_ma_btn
			CancelButton 555, y_pos, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation

		err_msg = "LOOP"

		If ButtonPressed = -1 Then ButtonPressed = next_page_ma_btn

		If all_questions_4_7_blank = "Questions 4 - 7 are not required." Then
			apply_for_ma = "Not Required"
			ma_self_employed = "Not Required"
			ma_start_working = "Not Required"
			ma_other_income = "Not Required"
		End If

		If ButtonPressed = add_memb_btn Then
			If NEW_MA_REQUEST_ARRAY(ma_request_client, 0) = "Select or Type" Then
				NEW_MA_REQUEST_ARRAY(ma_request_client, 0) = "Enter or Select Member"
			Else
				new_item = UBound(NEW_MA_REQUEST_ARRAY, 2) + 1
				ReDim Preserve NEW_MA_REQUEST_ARRAY(ma_request_notes, new_item)
				NEW_MA_REQUEST_ARRAY(ma_request_client, new_item) = "Enter or Select Member"
			End If
		End If
		If ButtonPressed = add_busi_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_type, new_earned_counter) = "BUSI"
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "MA"
			new_earned_counter = new_earned_counter + 1
		End If
		If ButtonPressed = add_jobs_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_type, new_earned_counter) = "JOBS"
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "MA"
			new_earned_counter = new_earned_counter + 1
		End If
		If ButtonPressed = add_unea_btn Then
			new_item = UBound(NEW_UNEARNED_ARRAY, 2) + 1
			ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, new_unearned_counter)
			NEW_UNEARNED_ARRAY(unearned_type, new_unearned_counter) = "UNEA"
			NEW_UNEARNED_ARRAY(unearned_prog_list, new_unearned_counter) = "MA"
			new_unearned_counter = new_unearned_counter + 1
		End If

		If ButtonPressed = next_page_ma_btn Then
			questions_answered = TRUE
			err_msg = ""

			If apply_for_ma = "Select One..." Then questions_answered = FALSE
			If ma_self_employed = "Select One..." Then questions_answered = FALSE
			If ma_start_working = "Select One..." Then questions_answered = FALSE
			If ma_other_income = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."
				If apply_for_ma = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 4 about applying for someone not currently getting MA coverage."
				If ma_self_employed = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 5 about anyone being self-employed."
				If ma_start_working = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 6 about anyone working."
				If ma_other_income = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 7 about unearned income."
			End If

			q_4_details_entered = FALSE
			q_5_details_entered = FALSE
			q_6_details_entered = FALSE
			q_7_details_entered = FALSE
			If InStr(apply_for_ma, "details listed below") <> 0 Then
				For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
					If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Enter or Select Member" Then
						q_4_details_entered = TRUE
					End If
				Next
			Else
				q_4_details_entered = TRUE
			End If
			If InStr(ma_self_employed, "details listed below") <> 0 Then
				For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
						q_5_details_entered = TRUE
					End If
				Next
			Else
				q_5_details_entered = TRUE
			End If
			If InStr(ma_start_working, "details listed below") <> 0 Then
				For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
						q_6_details_entered = TRUE
					End If
				Next
			Else
				q_6_details_entered = TRUE
			End If
			If InStr(ma_other_income, "details listed below") <> 0 Then
				For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
					If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
						q_7_details_entered = TRUE
					End If
				Next
			Else
				q_7_details_entered = TRUE
			End If

			If q_4_details_entered = FALSE OR q_5_details_entered = FALSE OR q_6_details_entered = FALSE  OR q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
			If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 4 - Someone getting MA coverage."
			If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 5 - Self-Employment Income."
			If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 6 - Job Income. "
			If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 7 - Unearned Income."
			If q_4_details_entered = FALSE OR q_5_details_entered = FALSE OR q_6_details_entered = FALSE  OR q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"

			If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
			If err_msg = "" Then csr_dlg_q_4_7_cleared = TRUE
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_4_7 = FALSE

end function

function csr_dlg_q_9_12()
	Do
		dlg_len = 205
		q_9_grp_len = 30
		q_10_grp_len = 30
		q_11_grp_len = 30
		q_12_grp_len = 30
		For new_assets_listed = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "CASH" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_9_grp_len = q_9_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "ACCT" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_9_grp_len = q_9_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_10_grp_len = q_10_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_11_grp_len = q_11_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_12_grp_len = q_12_grp_len + 20
			End If
		Next
		y_pos = 25
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, dlg_len, "MA CSR Asset Questions"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 9 - 12:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 9 - 12 are not required.", all_questions_9_12_blank

		  GroupBox 15, y_pos + 5, 585, q_9_grp_len, "Q9. Does anyone have cash, a savings or checking account, or a certificate of deposit?"
		  DropListBox 330, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_liquid_assets
		  y_pos = y_pos +20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_acct_btn
		  first_account = TRUE

		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			  If first_account = TRUE Then
				  Text 30, y_pos, 55, 10, "Owner(s) Name"
				  Text 165, y_pos, 25, 10, "Type"
				  Text 285, y_pos, 50, 10, "Bank Name"
				  y_pos = y_pos + 10
				  first_account = FALSE
			  End If
			  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset) 'liquid_asset_member'
			  ComboBox 165, y_pos, 115, 40, "Type or Select..."+chr(9)+"Cash"+chr(9)+ACCT_type_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)'liquid_asst_type
			  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)'liquid_asset_name
			  y_pos = y_pos + 20
			End If
		  Next
		  If first_account = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no ACCT information has been added."
			y_pos = y_pos + 10
		  End If

		  y_pos = y_pos +10
		  GroupBox 15, y_pos + 5, 585, q_10_grp_len, "Q10. Does anyone own or co-own securities or other assets?"
		  DropListBox 295, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_security_assets
		  y_pos = y_pos +  20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_secu_btn

		  first_secu = TRUE
		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				If first_secu = TRUE Then
					Text 30, y_pos, 55, 10, "Owner(s) Name"
					Text 165, y_pos, 25, 10, "Type"
					Text 285, y_pos, 50, 10, "Bank Name"
					y_pos = y_pos + 10
					first_secu = FALSE
				End If
				ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset) 'security_asset_member
				ComboBox 165, y_pos, 115, 40, "Type or Select"+chr(9)+SECU_type_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'security_asset_type
				EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)   'security_asset_name
				y_pos = y_pos + 20
			End If
		  Next
		  If first_secu = TRUE Then
			  Text 30, y_pos, 250, 10, "CSR form - no SECU information has been added."
			  y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_11_grp_len, "Q11. Does anyone own a vehicle?"
		  DropListBox 250, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_vehicle
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_cars_btn
		  first_car = TRUE
		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			  If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				  If first_car = TRUE Then
					  Text 30, y_pos, 55, 10, "Owner(s) Name"
					  Text 165, y_pos, 25, 10, "Type"
					  Text 285, y_pos, 75, 10, "Year/Make/Model"
					  y_pos = y_pos + 10
					  first_car = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset)     'vehicle_asset_member
				  ComboBox 165, y_pos, 115, 40, "Type or Select"+chr(9)+CARS_type_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'vehicle_asset_type
				  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_year_make_model, each_asset)  'vehicle_asset_name
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_car = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no CARS information has been added."
			y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_12_grp_len, "Q12. Does anyone own or co-own any real estate?"
		  DropListBox 280, y_pos, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_real_assets
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_rest_btn
		  first_home = TRUE
		  For each_asset = 0 to Ubound(NEW_ASSET_ARRAY, 2)
			  If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				  If first_home = TRUE Then
					  Text 30, y_pos, 55, 10, "Owner(s) Name"
					  Text 165, y_pos, 25, 10, "Address"
					  Text 320, y_pos, 75, 10, "Type of Property"
					  y_pos = y_pos + 10
					  first_home = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset)     'property_asset_member
				  EditBox 165, y_pos, 150, 15, NEW_ASSET_ARRAY(asset_address, each_asset)      'property_asset_address
				  ComboBox 320, y_pos, 150, 40, "Type or Select"+chr(9)+REST_type_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)     'property_asset_type
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_home = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no REST information has been added."
			y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  ButtonGroup ButtonPressed
			PushButton 415, y_pos, 80, 15, "Go Back to Q4 - Q7", back_to_ma_dlg_1
			PushButton 495, y_pos, 60, 15, "Continue", continue_btn
			CancelButton 555, y_pos, 50, 15
		EndDialog

		err_msg = "LOOP"

		dialog Dialog1
		cancel_confirmation

		If ButtonPressed = -1 Then ButtonPressed = continue_btn

		If ButtonPressed = add_acct_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "ACCT"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_secu_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "SECU"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_cars_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "CARS"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_rest_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "REST"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If

		If all_questions_9_12_blank = "Questions 9 - 12 are not required." Then
			ma_liquid_assets = "Not Required"
			ma_security_assets = "Not Required"
			ma_vehicle = "Not Required"
			ma_real_assets = "Not Required"
		End If

		If ButtonPressed = continue_btn Then
			questions_answered = TRUE
			err_msg = ""

			If ma_liquid_assets = "Select One..." Then questions_answered = FALSE
			If ma_security_assets = "Select One..." Then questions_answered = FALSE
			If ma_vehicle = "Select One..." Then questions_answered = FALSE
			If ma_real_assets = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."

				If ma_liquid_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_security_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_vehicle = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_real_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
			End If

			q_9_details_entered = FALSE
			q_10_details_entered = FALSE
			q_11_details_entered = FALSE
			q_12_details_entered = FALSE
			If InStr(ma_liquid_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" AND NEW_ASSET_ARRAY(ma_request_client, each_asset) <> "Enter or Select Member" Then
						q_9_details_entered = TRUE
					End If
				Next
			Else
				q_9_details_entered = TRUE
			End If
			If InStr(ma_security_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_10_details_entered = TRUE
					End If
				Next
			Else
				q_10_details_entered = TRUE
			End If
			If InStr(ma_vehicle, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_11_details_entered = TRUE
					End If
				Next
			Else
				q_11_details_entered = TRUE
			End If
			If InStr(ma_real_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_12_details_entered = TRUE
					End If
				Next
			Else
				q_12_details_entered = TRUE
			End If

			If q_9_details_entered = FALSE OR q_10_details_entered = FALSE OR q_11_details_entered = FALSE  OR q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
			If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 9 - Liquid Assets."
			If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-  Question 10 - Securities."
			If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 11 - Vehicles."
			If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 12 - Real Estate."
			If q_9_details_entered = FALSE OR q_10_details_entered = FALSE OR q_11_details_entered = FALSE  OR q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"


			If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
			If err_msg = "" Then 	csr_dlg_q_9_12_cleared = TRUE
		End If

		If ButtonPressed = back_to_ma_dlg_1 Then
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_13 = FALSE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
			err_msg = ""
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_9_12 = FALSE

end function

function csr_dlg_q_13()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, 80, "MA CSR Changes"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 13:"

		  Text 25, 25, 135, 10, "Q13. Do you have any changes to report?"
		  DropListBox 160, 20, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_other_changes
		  EditBox 30, 40, 555, 15, other_changes_reported

		  ButtonGroup ButtonPressed
			PushButton 255, 60, 100, 15, "Back to Q 4-7", back_to_ma_dlg_1
			PushButton 355, 60, 100, 15, "Back to Q 9 - 12", back_to_ma_dlg_2
			PushButton 455, 60, 100, 15, "Finish MA Questions", finish_ma_questions
			CancelButton 555, 60, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation
		If ButtonPressed = -1 Then ButtonPressed = finish_ma_questions

		If ButtonPressed = back_to_ma_dlg_1 Then
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
		End If
		If ButtonPressed = back_to_ma_dlg_2 Then
			show_csr_dlg_q_9_12 = TRUE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
		End If
		If ButtonPressed = finish_ma_questions Then
			show_ma_dlg_three = FALSE

			questions_answered = TRUE

			If trim(other_changes_reported) <> "" Then details_shown = TRUE
			If ma_other_changes = "Not Required" Then details_shown = TRUE

			If ma_other_changes = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."

				If ma_other_changes = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Indicate what the client entered for Question 13."
			Else
				If details_shown = FALSE Then err_msg = err_msg & vbNewLine & "* You must either enter what the client wrote in for Question 13 or check the box to indicate if if was blank."
			End If
			If trim(other_changes_reported) <> "" AND changes_reported_blank_checkbox = checked Then err_msg = err_msg & vbNewLine & "* You entered detail in what the client wrote and indicated it was blank using the checkbox, please update as only one of these should be completed."

			If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
			If err_msg = "" Then csr_dlg_q_13_cleared = TRUE
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_13 = FALSE
end function

function csr_dlg_q_15_19()
	Do
		err_msg = ""

		dlg_len = 190
		q_15_grp_len = 30
		q_16_grp_len = 25
		q_17_grp_len = 25
		q_18_grp_len = 25

		For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
			If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
				dlg_len = dlg_len + 20
				q_16_grp_len = q_16_grp_len + 20
			End If
		Next

		For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
				dlg_len = dlg_len + 20
				q_17_grp_len = q_17_grp_len + 20
			End If
		Next

		For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
			If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
				dlg_len = dlg_len + 20
				q_18_grp_len = q_18_grp_len + 20
			End If
		Next

		y_pos = 95

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 615, dlg_len, "SNAP CSR Question Details"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 15 - 19:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 15 - 19 are not required.", all_questions_15_19_blank
		  GroupBox 10, 30, 600, q_15_grp_len, "Q15. Has your household moved since your last application or in the past six months?"
		  DropListBox 305, 25, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_fifteen_form_answer
		  Text 25, 45, 105, 10, "New Rent or Mortgage Amount:"
		  EditBox 130, 40, 65, 15, new_rent_or_mortgage_amount
		  CheckBox 220, 45, 50, 10, "Heat/AC", heat_ac_checkbox
		  CheckBox 275, 45, 50, 10, "Electricity", electricity_checkbox
		  CheckBox 345, 45, 50, 10, "Telephone", telephone_checkbox
		  GroupBox 10, 70, 490, q_16_grp_len, "Q16 Has there been a change in EARNED INCOME?"
		  DropListBox 190, 65, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_sixteen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, 70, 50, 10, "Add Another", add_snap_earned_income_btn
		  first_earned = TRUE
		  For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
				  If first_earned = TRUE Then
					  Text 15, 85, 20, 10, "Client"
					  Text 130, 85, 100, 10, "Employer (or Business Name)"
					  Text 265, 85, 50, 10, "Change Date"
					  Text 320, 85, 35, 10, "Amount"
					  Text 375, 85, 40, 10, "Frequency"
					  Text 445, 85, 25, 10, "Hours"
					  first_earned = FALSE
				  End If
				  ComboBox 15, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, the_earned)
				  EditBox 130, y_pos, 130, 15, NEW_EARNED_ARRAY(earned_source, the_earned)
				  EditBox 265, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_change_date, the_earned)
				  EditBox 320, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_amount, the_earned)
				  DropListBox 375, y_pos, 65, 45, "Select One..."+chr(9)+"Weekly"+chr(9)+"BiWeekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", NEW_EARNED_ARRAY(earned_freq, the_earned)
				  EditBox 445, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_hours, the_earned)
				  y_pos = y_pos + 20
			  End If
		  Next
		  y_pos = y_pos + 10
		  GroupBox 10, y_pos, 490, q_17_grp_len, "Q17. Has there been a change in UNEARNED INCOME?"
		  DropListBox 205, y_pos - 5, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_seventeen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, y_pos, 50, 10, "Add Another", add_snap_unearned_btn
		  y_pos = y_pos + 15
		  first_unearned = TRUE
		  For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			  If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
				  If first_unearned = TRUE Then
					  Text 15, y_pos, 20, 10, "Client"
					  Text 145, y_pos, 100, 10, "Type and Source"
					  Text 280, y_pos, 50, 10, "Change Date"
					  Text 340, y_pos, 35, 10, "Amount"
					  Text 405, y_pos, 40, 10, "Frequency"
					  y_pos = y_pos + 10
					  first_unearned = FALSE
				  End If
				  ComboBox 15, y_pos, 125, 45, all_the_clients, NEW_UNEARNED_ARRAY(unearned_client, the_unearned)
				  EditBox 145, y_pos, 130, 15, NEW_UNEARNED_ARRAY(unearned_source, the_unearned)
				  EditBox 280, y_pos, 55, 15, NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)
				  EditBox 340, y_pos, 60, 15, NEW_UNEARNED_ARRAY(unearned_amount, the_unearned)
				  DropListBox 405, y_pos, 90, 45, "Select One..."+chr(9)+"Weekly"+chr(9)+"BiWeekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", NEW_UNEARNED_ARRAY(unearned_freq, the_unearned)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_unearned = TRUE Then y_pos = y_pos + 10
		  y_pos = y_pos + 10
		  GroupBox 10, y_pos, 490, q_18_grp_len, "Q18 Has there been a change in CHILD SUPPORT?"
		  DropListBox 190, y_pos - 5, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_eighteen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, y_pos, 50, 10, "Add Another", add_snap_cs_btn
		  y_pos = y_pos + 15

		  first_cs = TRUE
		  For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
			  If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
				  If first_cs = TRUE Then
					  Text 15, y_pos, 85, 10, "Name of person paying"
					  Text 220, y_pos, 35, 10, "Amount"
					  Text 295, y_pos, 65, 10, "Currently Paying?"
					  y_pos = y_pos + 10
					  first_cs = FALSE
				  End If
				  EditBox 15, y_pos, 200, 15, NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs)
				  EditBox 220, y_pos, 65, 15, NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)
				  DropListBox 295, y_pos, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_cs = TRUE Then y_pos = y_pos + 10
		  y_pos = y_pos + 10
		  Text 10, y_pos, 345, 10, "Q19. Did you work 20 hours each week, for an average of 80 hours per month during the past six months?"
		  DropListBox 355, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Not Required", quest_nineteen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 505, y_pos-5, 50, 15, "Continue", continue_btn
			CancelButton 555, y_pos-5, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If all_questions_15_19_blank = "Questions 15 - 19 are not required." Then
			quest_fifteen_form_answer = "Not Required"
			quest_sixteen_form_answer = "Not Required"
			quest_seventeen_form_answer = "Not Required"
			quest_eighteen_form_answer = "Not Required"
			quest_nineteen_form_answer = "Not Required"
		End If

		If quest_fifteen_form_answer = "Select One..." OR quest_sixteen_form_answer = "Select One..." OR quest_seventeen_form_answer = "Select One..." OR quest_eighteen_form_answer = "Select One..." OR quest_nineteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form." & vbNewLine
		If quest_fifteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 15 (Has the household moved?)."
		If quest_sixteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 16 (Has anyone had a change in Earned income?)."
		If quest_seventeen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 17 (Has anyone had a change in Unearned income?)."
		If quest_eighteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 18 (Has there been a change in Child Support income?)."
		If quest_nineteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 19 (Have you worked 80 hours per month?)."

		q_15_details_entered = FALSE
		q_16_details_entered = FALSE
		q_17_details_entered = FALSE
		q_18_details_entered = FALSE

		If InStr(quest_fifteen_form_answer, "details listed below") <> 0 Then
			new_rent_or_mortgage_amount = trim(new_rent_or_mortgage_amount)
			If new_rent_or_mortgage_amount <> "" Then q_15_details_entered = TRUE
			If heat_ac_checkbox = CHECKED Then q_15_details_entered = TRUE
			If electricity_checkbox = CHECKED Then q_15_details_entered = TRUE
			If telephone_checkbox = CHECKED Then q_15_details_entered = TRUE
		Else
			q_15_details_entered = TRUE
		End If

		If InStr(quest_sixteen_form_answer, "details listed below") <> 0 Then
			For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
				If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
					q_16_details_entered = TRUE
				End If
			Next
		Else
			q_16_details_entered = TRUE
		End If

		If InStr(quest_seventeen_form_answer, "details listed below") <> 0 Then
			For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
				If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
					q_17_details_entered = TRUE
				End If
			Next
		Else
			q_17_details_entered = TRUE
		End If

		If InStr(quest_eighteen_form_answer, "details listed below") <> 0 Then
			For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
				If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
					q_18_details_entered = TRUE
				End If
			Next
		Else
			q_18_details_entered = TRUE
		End If

		If q_15_details_entered = FALSE OR q_16_details_entered = FALSE OR q_17_details_entered = FALSE  OR q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
		If q_15_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 15 - Shelter and Utilities Expenses."
		If q_16_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 16 - Earned Income."
		If q_17_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 17 - Unearned Income."
		If q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 18 - Child Support."
		If q_15_details_entered = FALSE OR q_16_details_entered = FALSE OR q_17_details_entered = FALSE  OR q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"


		If ButtonPressed = add_snap_earned_income_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "SNAP"
			new_earned_counter = new_earned_counter + 1
			err_msg = "LOOP" & err_msg
		End If

		If ButtonPressed = add_snap_unearned_btn Then
			new_item = UBound(NEW_UNEARNED_ARRAY, 2) + 1
			ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, new_unearned_counter)
			NEW_UNEARNED_ARRAY(unearned_prog_list, new_unearned_counter) = "SNAP"
			new_unearned_counter = new_unearned_counter + 1
			err_msg = "LOOP" & err_msg
		End If

		If ButtonPressed = add_snap_cs_btn Then
			If NEW_CHILD_SUPPORT_ARRAY(cs_current, 0) = "" THen
				NEW_CHILD_SUPPORT_ARRAY(cs_current, 0) = "Select..."
			Else
				new_item = UBound(NEW_CHILD_SUPPORT_ARRAY, 2) + 1
				ReDim Preserve NEW_CHILD_SUPPORT_ARRAY(cs_notes, new_item)
			End If
			err_msg = "LOOP" & err_msg

		End If

		If ButtonPressed = -1 Then ButtonPressed = continue_btn

		If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = ""
	show_csr_dlg_q_15_19 = FALSE
	csr_dlg_q_15_19_cleared = TRUE
end function

function csr_dlg_sig()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 201, 95, "Form dates and signatures"
		  EditBox 135, 35, 60, 15, csr_form_date
		  DropListBox 135, 55, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_yn
		  ButtonGroup ButtonPressed
		    PushButton 35, 75, 105, 15, "Complete CSR Form Detail", complete_csr_questions
		    CancelButton 145, 75, 50, 15
		  Text 70, 40, 55, 10, "CSR Form Date:"
		  Text 10, 60, 120, 10, "Cient signature accepted verbally?"
		  Text 10, 10, 160, 20, "Confirm the client is signing this form and attesting to the information provided verbally."
		EndDialog

		dialog Dialog1

		cancel_confirmation

		If IsDate(csr_form_date) = FALSE Then
			err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was received."
		Else
			If DateDiff("d", date, csr_form_date) > 0 Then err_msg = err_msg & vbNewLine & "* The date of the CSR form is listed as a future date, a form cannot be listed as received inthe future, please review the form date."
		End If
		If client_signed_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the client has signed the form correctly by selecting 'yes' or 'no'."

		If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg

	Loop until err_msg = ""
	show_csr_dlg_sig = FALSE
	csr_dlg_sig_cleared = TRUE
end function

function confirm_csr_form_dlg()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 751, 370, "CSR Form Information"
	  If show_buttons_on_confirmation_dlg = TRUE Then
		  DropListBox 255, 350, 190, 50, "Indicate the form information"+chr(9)+"NO - the information here is different"+chr(9)+"YES - This is the information on the CSR Form", confirm_csr_form_information
		  ButtonGroup ButtonPressed
		    OkButton 645, 350, 50, 15
		    CancelButton 695, 350, 50, 15
			PushButton 15, 327, 165, 13, "Fix Page One Information", back_to_dlg_addr
			PushButton 200, 327, 165, 13, "Fix Page Two Information", back_to_dlg_ma_income
			PushButton 385, 327, 165, 13, "Fix Page Three Information", back_to_dlg_ma_asset
			PushButton 570, 270, 165, 13, "Fix Page Four Information", back_to_dlg_snap
			PushButton 570, 327, 165, 13, "Fix Page Five Information", back_to_dlg_sig
	  Else
		  ButtonGroup ButtonPressed
			OkButton 695, 350, 50, 15
	  End If
	  GroupBox 5, 5, 185, 340, "Page 1"
	  Text 10, 20, 105, 10, "1. Name and Address"
	  Text 20, 35, 160, 10, "Name:" & client_on_csr_form
	  Text 20, 50, 70, 10, "Residence Address"
	  If new_resi_addr_entered = TRUE Then
		  Text 25, 65, 110, 10, new_resi_one
		  Text 25, 75, 110, 10, new_resi_city & ", " & new_resi_state & " " & new_resi_zip
		  y_pos_1 = 85
	  Else
		  Text 25, 65, 110, 10, resi_line_one
		  If resi_line_two = "" Then
			Text 25, 75, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
			y_pos_1 = 85
		  Else
			Text 25, 75, 110, 10, resi_line_two
			Text 25, 85, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
			y_pos_1 = 95
		  End If
	  End If
	  If residence_address_match_yn = "Yes - the addresses are the same." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - YES"
	  If residence_address_match_yn = "No - there is a difference." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If residence_address_match_yn = "No - New Address Entered" Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If residence_address_match_yn = "RESI Address not Provided" Then Text 100, y_pos_1, 75, 10, "BLANK RESI ADDR"

	  Text 20, y_pos_1 + 15, 70, 10, "Mailing Address"
	  y_pos_1 = y_pos_1 + 30
	  If new_mail_addr_entered = TRUE Then
		  Text 25, y_pos_1, 110, 10, new_mail_one
		  y_pos_1 = y_pos_1 + 10
		  Text 25, y_pos_1, 110, 10, new_mail_city & ", " & new_mail_state & " " & new_mail_zip
		  y_pos_1 = y_pos_1 + 10
	  Else
		  If mail_line_one = "" Then
			  Text 25, y_pos_1, 110, 10, "NO MAILING ADDRESS LISTED"
			  y_pos_1 = y_pos_1 + 15
		  Else
			  Text 25, y_pos_1, 110, 10, mail_line_one
			  y_pos_1 = y_pos_1 + 10
			  If mail_line_two = "" Then
				Text 25, y_pos_1, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
				y_pos_1 = y_pos_1 + 10
			  Else
				Text 25, y_pos_1, 110, 10, mail_line_two
				Text 25, y_pos_1 + 10, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
				y_pos_1 = y_pos_1 + 20
			  End If
		  End If
	  End If
	  If mailing_address_match_yn = "Yes - the addresses are the same." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - YES"
	  If mailing_address_match_yn = "No - there is a difference." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If mailing_address_match_yn = "No - New Address Entered" Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If mailing_address_match_yn = "MAIL Address not Provided" Then Text 100, y_pos_1, 75, 10, "BLANK MAIL ADDR"

	  y_pos_1 = y_pos_1 + 15
	  Text 10, y_pos_1, 110, 20, "2. Has anyone moved in or out of your home in the past six months?"
	  Text 150, y_pos_1, 35, 10, quest_two_move_in_out
	  y_pos_1 = y_pos_1 + 25
	  pers_list_count = 1

	  For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		  If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked OR ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then
			  Text 20, y_pos_1, 50, 10, "Person 1" & pers_list_count
			  If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then Text 140, y_pos_1, 40, 10, "MOVED OUT"
			  If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then Text 140, y_pos_1, 40, 10, "MOVED IN"
			  Text 25, y_pos_1 + 10, 155, 10, "Name:" & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb)
			  If len(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb)) > 4 Then Text 25, y_pos_1 + 20, 155, 10, "Relationship:" & right(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb), len(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb)) - 5)
			  pers_list_count = pers_list_count + 1
			  y_pos_1 = y_pos_1 + 40
		  End If
	  Next
	  For new_memb_counter = 0 to UBOUND(NEW_MEMBERS_ARRAY, 2)
		  If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked OR NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then
			  Text 20, y_pos_1, 50, 10, "Person 1" & pers_list_count
			  If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked Then Text 140, y_pos_1, 40, 10, "MOVED OUT"
			  If NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then Text 140, y_pos_1, 40, 10, "MOVED IN"
			  Text 25, y_pos_1 + 10, 155, 10, "Name:" & NEW_MEMBERS_ARRAY(new_first_name, new_memb_counter) & " " & NEW_MEMBERS_ARRAY(new_last_name, new_memb_counter)
			  Text 25, y_pos_1 + 20, 155, 10, "Relationship:" & right(NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter), len(NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter)) - 5)
			  pers_list_count = pers_list_count + 1
			  y_pos_1 = y_pos_1 + 40
		  End If
	  Next

	  GroupBox 190, 5, 185, 340, "Page 2"
	  Text 195, 20, 135, 20, "4. Do you want to apply for someone who is not getting coverage now?"
	  Text 340, 20, 35, 10, replace(apply_for_ma, "Not Required", "BLANK")
	  y_pos_2 = 40
	  If q_4_details_blank_checkbox = checked then
		  Text 200, y_pos_2, 150, 10, "Q4 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
		  If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
			  Text 200, y_pos_2, 150, 10, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 135, 20, "5. Is anyone self-employed or does anyone expect to be self-employed?"
	  Text 340, y_pos_2, 35, 10, replace(ma_self_employed, "Not Required", "BLANK")
	  y_pos_2 = y_pos_2 + 20
	  If q_5_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q5 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_EARNED_ARRAY(earned_client, each_busi) & " from " & NEW_EARNED_ARRAY(earned_source, each_busi) & " - $" & NEW_EARNED_ARRAY(earned_amount, each_busi) & " on " & NEW_EARNED_ARRAY(earned_start_date, each_busi)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 135, 20, "6. Does anyone work or does anyone expect to start working?"
	  Text 340, y_pos_2, 35, 10, replace(ma_start_working, "Not Required", "BLANK")
	  y_pos_2 = y_pos_2 + 20
	  If q_6_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q6 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_EARNED_ARRAY(earned_client, each_job) & " from " & NEW_EARNED_ARRAY(earned_source, each_job) & " - $" & NEW_EARNED_ARRAY(earned_amount, each_job) & " on " & NEW_EARNED_ARRAY(earned_start_date, each_job)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 140, 25, "7. Does anyone get money or does anyone expect to get money from sources other than work?"
	  Text 340, y_pos_2, 35, 10, replace(ma_other_income, "Not Required", "BLANK")
	  y_pos_2 = y_pos_2 + 30
	  If q_7_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q7 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		  If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_UNEARNED_ARRAY(unearned_client, each_unea) & " from " & NEW_UNEARNED_ARRAY(unearned_source, each_unea) & " - $" & NEW_UNEARNED_ARRAY(unearned_amount, each_unea) & " on " & NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10


	  GroupBox 375, 5, 185, 340, "Page 3"
	  Text 380, 20, 145, 10, "9. Does anyone have cash or account?"
	  Text 525, 20, 35, 10, replace(ma_liquid_assets, "Not Required", "BLANK")
	  y_pos_3 = 30
	  If q_9_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q9 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		  Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
		  y_pos_3 = y_pos_3 + 10
		End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 135, 20, "10. Does anyone own securities or other assets?"
	  Text 525, y_pos_3, 35, 10, replace(ma_security_assets, "Not Required", "BLANK")
	  y_pos_3 = y_pos_3 + 20
	  If q_10_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q10 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			y_pos_3 = y_pos_3 + 10
		End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 135, 10, "11. Does anyone own a vehicle?"
	  Text 525, y_pos_3, 35, 10, replace(ma_vehicle, "Not Required", "BLANK")
	  y_pos_3 = y_pos_3 + 10
	  If q_11_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q11 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		  If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			  Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			  y_pos_3 = y_pos_3 + 10
		  End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 140, 20, "12. Does anyone own or co-own a house or any real estate?"
	  Text 525, y_pos_3, 35, 10, replace(ma_real_assets, "Not Required", "BLANK")
	  y_pos_3 = y_pos_3 + 20
	  If q_12_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q12 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If

	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 140, 10, "13. Do you have any change to report?"
	  Text 525, y_pos_3, 35, 10, replace(ma_other_changes, "Not Required", "BLANK")
	  y_pos_3 = y_pos_3 + 10
	  If changes_reported_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q13 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  If trim(other_changes_reported) <> "" Then
		  Text 385, y_pos_3, 150, 10, "Other changes: " & other_changes_reported
		  y_pos_3 = y_pos_3 + 10
	  ENd If

	  y_pos_3 = y_pos_3 + 10


	  GroupBox 560, 5, 185, 340, "Page 4"
	  Text 585, 20, 135, 20, "Since your last application or in the past six months..."
	  Text 565, 45, 125, 10, "15. Has your household moved?"
	  Text 710, 45, 35, 10, replace(quest_fifteen_form_answer, "Not Required", "BLANK")
	  y_pos_4 = 55
	  If trim(new_rent_or_mortgage_amount) = "" AND heat_ac_checkbox = unchecked AND electricity_checkbox = unchecked AND telephone_checkbox = unchecked Then
		  Text 570, y_pos_4, 150, 10, "Q15 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  Else
		  If trim(new_rent_or_mortgage_amount) = "" Then Text 570, y_pos_4, 150, 10, "NO new shelter Cost"
		  If trim(new_rent_or_mortgage_amount) <> "" THen Text 570, y_pos_4, 150, 10, "New Shelter Cost: $" & new_rent_or_mortgage_amount
		  y_pos_4 = y_pos_4 + 10

		  If heat_ac_checkbox = checked OR electricity_checkbox = checked OR telephone_checkbox = checked Then
			  Text 570, y_pos_4, 50, 10, "Utilities Paid"
			  y_pos_4 = y_pos_4 + 10
			  If heat_ac_checkbox = checked Then Text 575, y_pos_4, 50, 10, "HEAT/AC"
			  If electricity_checkbox = checked Then Text 625, y_pos_4, 50, 10, "ELECTRIC"
			  If telephone_checkbox = checked Then Text 675, y_pos_4, 50, 10, "PHONE"
			  y_pos_4 = y_pos_4 + 10
		  End If
	  End If

	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 20, "16. Has anyone had a change in their income from work?"
	  Text 710, y_pos_4, 35, 10, replace(quest_sixteen_form_answer, "Not Required", "BLANK")
	  y_pos_4 = y_pos_4 + 20
	  If q_16_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q16 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
			  Text 570, y_pos_4, 150, 10, NEW_EARNED_ARRAY(earned_client, the_earned) & " from " & NEW_EARNED_ARRAY(earned_source, the_earned) & " - $" & NEW_EARNED_ARRAY(earned_amount, the_earned) & " on " & NEW_EARNED_ARRAY(earned_change_date, the_earned)
			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 140, 25, "17. Has anyone had a change of more than $50 per month from income sources other than work or a change in unearned income?"
	  Text 710, y_pos_4, 35, 10, replace(quest_seventeen_form_answer, "Not Required", "BLANK")
	  y_pos_4 = y_pos_4 + 30
	  If q_17_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q17 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		  If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
			  Text 570, y_pos_4, 150, 10, NEW_UNEARNED_ARRAY(unearned_client, the_unearned) & " from " & NEW_UNEARNED_ARRAY(unearned_source, the_unearned) & " - $" & NEW_UNEARNED_ARRAY(unearned_amount, the_unearned) & " on " & NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)

			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 25, "18. Has anyone had a change in court-ordered child or medical support payments?"
	  Text 710, y_pos_4, 35, 10, replace(quest_eighteen_form_answer, "Not Required", "BLANK")
	  y_pos_4 = y_pos_4 + 30
	  If q_18_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q18 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
		  If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
			  Text 570, y_pos_4, 150, 10, "Child support - paid by: " & NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) & " - $" & NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)
			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 25, "19. Did you work 20 hours each week, for an average of 80 hours each month during the past six months?"
	  Text 710, y_pos_4, 35, 10, replace(quest_nineteen_form_answer, "Not Required", "BLANK")
	  GroupBox 560, 285, 185, 60, "Page 5"
	  Text 570, 300, 165, 10, "Signature:" & client_signed_yn
	  Text 570, 315, 165, 10, "Date:" & csr_form_date

	  Text 10, 355, 240, 10, "Review the information here, does it match the form the client submited?"

	EndDialog

	dialog Dialog1
	cancel_confirmation

	err_msg = "LOOP"

	If ButtonPressed = back_to_dlg_addr Then
		show_csr_dlg_q_1 = TRUE
		show_csr_dlg_q_2 = TRUE
	End If
	If ButtonPressed = back_to_dlg_ma_income Then show_csr_dlg_q_4_7 = TRUE
	If ButtonPressed = back_to_dlg_ma_asset Then
		show_csr_dlg_q_9_12 = TRUE
		show_csr_dlg_q_13 = TRUE
	End If
	If ButtonPressed = back_to_dlg_snap Then show_csr_dlg_q_15_19 = TRUE
	If ButtonPressed = back_to_dlg_sig Then show_csr_dlg_sig = TRUE
	If ButtonPressed = -1 Then
		err_msg = ""
		If confirm_csr_form_information = "Indicate the form information" THen err_msg = err_msg & vbNewLine & "* Indicate if this information is correct and matches the form received. If something is not correct, use the buttons on this dialog to go back to the correct area and update the information on the specific dialog."
		If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg
		If confirm_csr_form_information = "NO - the information here is different" Then
			show_csr_dlg_q_1 = TRUE
			show_csr_dlg_q_2 = TRUE
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_9_12 = TRUE
			show_csr_dlg_q_13 = TRUE
			show_csr_dlg_q_15_19 = TRUE
			show_csr_dlg_sig = TRUE
		End If
	Else
		confirm_csr_form_information = "Indicate the form information"
	End If
end function

function gather_pers_detail()
	If grh_sr_yn = "Yes" Then grh_sr = TRUE
	If hc_sr_yn = "Yes" Then hc_sr = TRUE
	If snap_sr_yn = "Yes" Then snap_sr = TRUE

	If grh_sr = TRUE Then
	    MAXIS_footer_month = grh_sr_mo
	    MAXIS_footer_year = grh_sr_yr
	End If
	If hc_sr = TRUE Then
	    MAXIS_footer_month = hc_sr_mo
	    MAXIS_footer_year = hc_sr_yr
	End If
	If snap_sr = TRUE Then
	    MAXIS_footer_month = snap_sr_mo
	    MAXIS_footer_year = snap_sr_yr
	End If

	Call navigate_to_MAXIS_screen("CASE", "PERS")

	pers_row = 10                                               'This is where client information starts on CASE PERS
	person_counter = 0
	Do
	    EMReadScreen the_snap_status, 1, pers_row, 54
	    EMReadScreen the_grh_status, 1, pers_row, 66
	    EMReadScreen the_hc_status, 1, pers_row, 61             'reading the HC status of each client
	    If the_snap_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Active"
	    ElseIf the_snap_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Inactive"
	    End If
	    If the_grh_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Active"
	    ElseIf the_grh_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Inactive"
	    End If
	    If the_hc_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Active"
	    ElseIf the_hc_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Inactive"
	    End If

	    person_counter = person_counter + 1
	    pers_row = pers_row + 3         'next client information is 3 rows down
	    If pers_row = 19 Then           'this is the end of the list of client on each list
	        PF8                         'going to the next page of client information
	        pers_row = 10
	        EmReadscreen end_of_list, 9, 24, 14
	        If end_of_list = "LAST PAGE" Then Exit Do
	    End If
	    EmReadscreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list

	Loop until next_pers_ref_numb = "  "
	Call back_to_SELF

	Call navigate_to_MAXIS_screen("STAT", "WREG")

	For all_the_membs = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
		If ALL_CLIENTS_ARRAY(clt_snap_status, all_the_membs) = "Active" OR ALL_CLIENTS_ARRAY(clt_snap_status, all_the_membs) = "Pending" Then
			EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, all_the_membs), 20, 76
			transmit

			EMReadScreen wreg_abawd_code, 2, 13, 50
			If wreg_abawd_code = "09" OR wreg_abawd_code = "10" OR wreg_abawd_code = "11" OR wreg_abawd_code = "13" Then abawd_on_case = TRUE
		End If
	Next
end function

function enter_new_residence_address()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 356, 135, "New Residence Address Information"
		  Text 240, 25, 50, 10, "Effective Date"
		  EditBox 300, 20, 50, 15, new_addr_effective_date
		  Text 10, 25, 145, 10, "New Residence Address Reported on CSR:"
		  Text 20, 45, 45, 10, "House/Street:"
		  EditBox 70, 40, 280, 15, new_resi_one
		  Text 50, 65, 15, 10, "City:"
		  EditBox 70, 60, 80, 15, new_resi_city
		  Text 160, 65, 20, 10, "State:"
		  DropListBox 185, 60, 75, 45, "Select One..."+chr(9)+state_list, new_resi_state
		  Text 275, 65, 20, 10, "Zip:"
		  EditBox 300, 60, 50, 15, new_resi_zip
		  Text 40, 85, 30, 10, "County:"
		  DropListBox 70, 80, 190, 45, "Select One..."+chr(9)+county_list, new_resi_county
		  Text 95, 100, 90, 10, "Address/Home Verification:"
		  DropListBox 190, 95, 125, 45, "Select One..."+chr(9)+"SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd", new_shel_verif
		  Text 10, 5, 300, 10, "ENTER THE RESIDENCE ADDRESS INFORMATION FROM THE CSR FORM"
		  ButtonGroup ButtonPressed
		    OkButton 300, 115, 50, 15
		EndDialog

		dialog Dialog1

		If trim(new_addr_effective_date) <> "" AND IsDate(new_addr_effective_date) = FALSE THen err_msg = err_msg & vbNewLine & "* Enter the effective date as a valid date or leave blank."
		new_resi_one = trim(new_resi_one)
		new_resi_city = trim(new_resi_city)
		new_resi_zip = trim(new_resi_zip)
		If new_resi_one = "" AND new_resi_city = "" AND new_resi_state = "Select One..." AND new_resi_zip = "" Then err_msg = err_msg & vbNewLine & "* Enter the details from the form."
		If new_resi_county = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the county of residence."
		If new_shel_verif = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the verification (NO or OT are acceptable)."

		If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg

	Loop until err_msg = ""
	residence_address_match_yn = "No - New Address Entered"
	new_resi_addr_entered = TRUE
end function

function enter_new_mailing_address()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 356, 95, "New Mailing Address Information"
		  Text 10, 20, 145, 10, "New Mailing Address Reported on CSR:"
		  Text 20, 40, 45, 10, "House/Street:"
		  EditBox 70, 35, 280, 15, new_mail_one
		  Text 50, 60, 15, 10, "City:"
		  EditBox 70, 55, 80, 15, new_mail_city
		  Text 160, 60, 20, 10, "State:"
		  DropListBox 185, 55, 75, 45, "Select One..."+chr(9)+state_list, new_mail_state
		  Text 275, 60, 20, 10, "Zip:"
		  EditBox 300, 55, 50, 15, new_mail_zip
		  Text 10, 5, 300, 10, "ENTER THE MAILING ADDRESS INFORMATION FROM THE CSR FORM"
		  ButtonGroup ButtonPressed
		    OkButton 300, 75, 50, 15
		EndDialog

		dialog Dialog1

		new_mail_one = trim(new_mail_one)
		new_mail_city = trim(new_mail_city)
		new_mail_zip = trim(new_mail_zip)

		If new_mail_one = "" AND new_mail_city = "" AND new_mail_state = "Select One..." AND new_mail_zip = "" Then err_msg = err_msg & vbNewLine & "* Enter the details from the form."
		If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg
	Loop until err_msg = ""
	mailing_address_match_yn = "No - New Address Entered"
	new_mail_addr_entered = TRUE
end function

function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
    If IsNumeric(footer_month) = FALSE Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be a number, review and reenter the footer month information."
    Else
        footer_month = footer_month * 1
        If footer_month > 12 OR footer_month < 1 Then err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be between 1 and 12. Review and reenter the footer month information."
        footer_month = right("00" & footer_month, 2)
    End If

    If len(footer_year) < 2 Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be at least 2 characters long, review and reenter the footer year information."
    Else
        If IsNumeric(footer_year) = FALSE Then
            err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be a number, review and reenter the footer year information."
        Else
            footer_year = right("00" & footer_year, 2)
        End If
    End If
end function

function save_your_work()
'This function will store the variables that have been defined in the script run into a text file.
'This file can then be accessed if the script fails or is cancelled.
'For scripts that work during 'live' interaction with clients/others this functionality is important to prevent loss of time and frustration
'This function is script specific because the variables need to be defined for each script within this dialog and the enumberation codes need to be script specific

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now creates the name of file
	local_work_save_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	'Needed for interacting with files in the computer
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		'If the file already exists, we need to delete it so that a brand new one can be written
		'If we do not delete it - the variables may be wrong/old.
		If .FileExists(local_work_save_path) = True then
			.DeleteFile(local_work_save_path)
		End If

		'At this point, the file should not exist in any case.
		'We need to create an write a new file.
		If .FileExists(local_work_save_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(local_work_save_path, ForWriting, true)

			'Write the contents of the text file - these are all the variables that save case information
			'Each element of every array is also saved here - joined together by a unique character so we can split them later.
			objTextStream.WriteLine "00 - " & all_the_clients
			objTextStream.WriteLine "FS SR - " & snap_sr_yn & " - " & snap_sr_mo & "/" & snap_sr_yr
			objTextStream.WriteLine "HC SR - " & hc_sr_yn & " - " & hc_sr_mo & "/" & hc_sr_yr
			objTextStream.WriteLine "GR SR - " & grh_sr_yn & " - " & grh_sr_mo & "/" & grh_sr_yr
			objTextStream.WriteLine "01 - " & client_on_csr_form
			objTextStream.WriteLine "01 - RESI - L1 - " & resi_line_one
			objTextStream.WriteLine "01 - RESI - L2 - " & resi_line_two
			objTextStream.WriteLine "01 - RESI - CI - " & resi_city
			objTextStream.WriteLine "01 - RESI - ST - " & resi_state
			objTextStream.WriteLine "01 - RESI - ZI - " & resi_zip
			objTextStream.WriteLine "01 - RESI - NEW - " & new_resi_addr_entered
			objTextStream.WriteLine "01 - RESI - NEW L1 - " & new_resi_one
			objTextStream.WriteLine "01 - RESI - NEW CI - " & new_resi_city
			objTextStream.WriteLine "01 - RESI - NEW ST - " & new_resi_state
			objTextStream.WriteLine "01 - RESI - NEW ZI - " & new_resi_zip

			objTextStream.WriteLine "01 - MAIL - L1 - " & mail_line_one
			objTextStream.WriteLine "01 - MAIL - L2 - " & mail_line_two
			objTextStream.WriteLine "01 - MAIL - CI - " & mail_city
			objTextStream.WriteLine "01 - MAIL - ST - " & mail_state
			objTextStream.WriteLine "01 - MAIL - ZI - " & mail_zip
			objTextStream.WriteLine "01 - MAIL - NEW - " & new_mail_addr_entered
			objTextStream.WriteLine "01 - MAIL - NEW L1 - " & new_mail_one
			objTextStream.WriteLine "01 - MAIL - NEW CI - " & new_mail_city
			objTextStream.WriteLine "01 - MAIL - NEW ST - " & new_mail_state
			objTextStream.WriteLine "01 - MAIL - NEW ZI - " & new_mail_zip

			objTextStream.WriteLine "01 - HMLS - " & homeless_status


			objTextStream.WriteLine "02 - " & quest_two_move_in_out
			objTextStream.WriteLine "02a- " & new_hh_memb_not_in_mx_yn
			For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
				objTextStream.WriteLine "ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & "~" & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & "~" & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & "~" & ALL_CLIENTS_ARRAY(memb_age, known_memb) & "~" & ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) & "~" & ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) & "~" & ALL_CLIENTS_ARRAY(clt_grh_status, known_memb) & "~" & ALL_CLIENTS_ARRAY(clt_hc_status, known_memb) & "~" & ALL_CLIENTS_ARRAY(clt_snap_status, known_memb)
				' ref~last_name~first_name~age~remo~new~grh~hc~snap
			Next
			For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
				objTextStream.WriteLine "NEW_MEMBERS_ARRAY - " & NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_dob, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_memb_moved_in, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_memb_moved_out, new_hh_memb) & "~" & NEW_MEMBERS_ARRAY(new_memb_notes, new_hh_memb)
				' first_name~mid_initial~last_name~suffix~dob~rel_to_applct~ma_req~fs~req~grh_req~moved_in~moved_out~notes
			Next
			objTextStream.WriteLine "04 - " & apply_for_ma
			For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
				If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
					objTextStream.WriteLine "NEW_MA_REQUEST_ARRAY - " & NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
				End If
			Next
			objTextStream.WriteLine "05 - " & ma_self_employed
			objTextStream.WriteLine "06 - " & ma_start_working
			objTextStream.WriteLine "07 - " & ma_other_income
			objTextStream.WriteLine "09 - " & ma_liquid_assets
			objTextStream.WriteLine "10 - " & ma_security_assets
			objTextStream.WriteLine "11 - " & ma_vehicle
			objTextStream.WriteLine "12 - " & ma_real_assets
			objTextStream.WriteLine "13 - " & ma_other_changes
			objTextStream.WriteLine "13 - DET - " & other_changes_reported
			objTextStream.WriteLine "15 - " & quest_fifteen_form_answer
			objTextStream.WriteLine "15 - RENT - " & new_rent_or_mortgage_amount
			objTextStream.WriteLine "15 - HEAT - " & heat_ac_checkbox
			objTextStream.WriteLine "15 - ELEC - " & electricity_checkbox
			objTextStream.WriteLine "15 - TELE - " & telephone_checkbox
			objTextStream.WriteLine "16 - " & quest_sixteen_form_answer
			objTextStream.WriteLine "17 - " & quest_seventeen_form_answer
			objTextStream.WriteLine "18 - " & quest_eighteen_form_answer
			objTextStream.WriteLine "19 - " & quest_nineteen_form_answer

			For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
				objTextStream.WriteLine "NEW_EARNED_ARRAY - " & NEW_EARNED_ARRAY(earned_client, each_job) & "~" & NEW_EARNED_ARRAY(earned_type, each_job) & "~" & NEW_EARNED_ARRAY(earned_source, each_job) & "~" & NEW_EARNED_ARRAY(earned_change_date, each_job) & "~" & NEW_EARNED_ARRAY(earned_amount, each_job) & "~" & NEW_EARNED_ARRAY(earned_freq, each_job) & "~" & NEW_EARNED_ARRAY(earned_hours, each_job) & "~" & NEW_EARNED_ARRAY(earned_prog_list, each_job) & "~" & NEW_EARNED_ARRAY(earned_start_date, each_job) & "~" & NEW_EARNED_ARRAY(earned_seasonal, each_job) & "~" & NEW_EARNED_ARRAY(earned_notes, each_job)
			Next

			For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
				objTextStream.WriteLine "NEW_UNEARNED_ARRAY - " & NEW_UNEARNED_ARRAY(unearned_client, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_type, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_source, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_change_date, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_amount, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_freq, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_start_date, each_unea) & "~" & NEW_UNEARNED_ARRAY(unearned_notes, each_unea)
			Next

			For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
				objTextStream.WriteLine "NEW_ASSET_ARRAY - " & NEW_ASSET_ARRAY(asset_client, each_asset) & "~" & NEW_ASSET_ARRAY(asset_type, each_asset) & "~" & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & "~" & NEW_ASSET_ARRAY(asset_bank_name, each_asset) & "~" & NEW_ASSET_ARRAY(asset_year_make_model, each_asset) & "~" & NEW_ASSET_ARRAY(asset_address, each_asset) & "~" & NEW_ASSET_ARRAY(asset_prog_list, each_asset) & "~" & NEW_ASSET_ARRAY(asset_notes, each_asset)
			Next
			For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
				objTextStream.WriteLine "NEW_CHILD_SUPPORT_ARRAY - " & NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) & "~" & NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs) & "~" & NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) & "~" & NEW_CHILD_SUPPORT_ARRAY(cs_notes, the_cs)
			Next

			'Close the object so it can be opened again shortly - these types of files do not need to be explicitly saved.
			objTextStream.Close
		End if
	End with
end function


function restore_your_work(vars_filled)
'This function runs at the beginning of the script after the case number is Gathered
'It will redefine the variables from the text file that was saved from a previous run if one exists.
'This function must be script specific as the variables are hard coded within in.

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Creates the file name based on the convention we have defined
	local_work_save_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	'Needed for interacting with files in the computer
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		'If the file exists it means that this script was run for this case by this worker.
		'If that happened, it may be that there was a failure that the worker will want to restore the information that was already saved.
		If .FileExists(local_work_save_path) = True then

			'Asking the worker if they would like to restore the information from the previous run
			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			'If the worker indicates 'YES' then we will read the file and save the information in the file BACK into the script variables.
			If pull_variables = vbYes Then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(local_work_save_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_csr_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE		'this tells the script that we do not need to read from MAXIS for ADDR and MEMB information

				'Setting initial counters for all the different arrays
				known_memb = 0
				new_hh_memb = 0
				each_new_memb = 0
				each_job = 0
				each_unea = 0
				each_asset = 0
				the_cs = 0
				'Now we look at each line in the text file and use the first characters to identify which variable the line is storing
				For Each text_line in saved_csr_details
					If left(text_line, 2) = "FS" Then
						answer = mid(text_line, 9, 3)
						answer = replace(answer, "-", "")
						snap_sr_yn = trim(answer)
						If right(text_line, 1) ="/" Then
							snap_sr_mo = ""
							snap_sr_yr = ""
						Else
							mo_yr = right(text_line, 5)
							snap_sr_mo = left(mo_yr, 2)
							snap_sr_yr = right(mo_yr, 2)
						End If
					End If
					If left(text_line, 2) = "HC" Then
						answer = mid(text_line, 9, 3)
						answer = replace(answer, "-", "")
						hc_sr_yn = trim(answer)
						If right(text_line, 1) ="/" Then
							hc_sr_mo = ""
							hc_sr_yr = ""
						Else
							mo_yr = right(text_line, 5)
							hc_sr_mo = left(mo_yr, 2)
							hc_sr_yr = right(mo_yr, 2)
						End If
					End If
					If left(text_line, 2) = "GR" Then
						answer = mid(text_line, 9, 3)
						answer = replace(answer, "-", "")
						grh_sr_yn = trim(answer)
						If right(text_line, 1) ="/" Then
							grh_sr_mo = ""
							grh_sr_yr = ""
						Else
							mo_yr = right(text_line, 5)
							grh_sr_mo = left(mo_yr, 2)
							grh_sr_yr = right(mo_yr, 2)
						End If
					End If
					If left(text_line, 2) = "00" Then
						all_the_clients = Mid(text_line, 6)
						list_for_array = right(all_the_clients, len(all_the_clients) - 15)
						full_hh_list = Split(list_for_array, chr(9))
					End If
					If left(text_line, 2) = "01" Then
						If mid(text_line, 6, 13) = "RESI - NEW L1" Then
							new_resi_one = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "RESI - NEW CI" Then
							new_resi_city = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "RESI - NEW ST" Then
							new_resi_state = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "RESI - NEW ZI" Then
							new_resi_zip = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 10) = "RESI - NEW" Then
							new_resi_addr_entered = Mid(text_line, 19)
						ElseIf mid(text_line, 6, 9) = "RESI - L1" Then
							resi_line_one = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "RESI - L2" Then
							resi_line_two = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "RESI - CI" Then
							resi_city = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "RESI - ST" Then
							resi_state = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "RESI - ZI" Then
							resi_zip = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 13) = "MAIL - NEW L1" Then
							new_mail_one = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "MAIL - NEW CI" Then
							new_mail_city = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "MAIL - NEW ST" Then
							new_mail_state = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 13) = "MAIL - NEW ZI" Then
							new_mail_zip = Mid(text_line, 22)
						ElseIf mid(text_line, 6, 10) = "MAIL - NEW" Then
							new_mail_addr_entered = Mid(text_line, 19)
						ElseIf mid(text_line, 6, 9) = "MAIL - L1" Then
							mail_line_one = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "MAIL - L2" Then
							mail_line_two = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "MAIL - CI" Then
							mail_city = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "MAIL - ST" Then
							mail_state = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 9) = "MAIL - ZI" Then
							mail_zip = Mid(text_line, 18)
						ElseIf mid(text_line, 6, 4) = "HMLS" Then
							homeless_status = Mid(text_line, 13)
						Else
							client_on_csr_form = Mid(text_line, 6)
						End If
					End If
					If left(text_line, 3) = "02 " Then quest_two_move_in_out = Mid(text_line, 6)
					If left(text_line, 3) = "02a" Then new_hh_memb_not_in_mx_yn = Mid(text_line, 6)

					If left(text_line, 2) = "04" Then apply_for_ma = Mid(text_line, 6)

					If left(text_line, 2) = "05" Then ma_self_employed = Mid(text_line, 6)

					If left(text_line, 2) = "06" Then ma_start_working = Mid(text_line, 6)

					If left(text_line, 2) = "07" Then ma_other_income = Mid(text_line, 6)

					If left(text_line, 2) = "09" Then ma_liquid_assets = Mid(text_line, 6)

					If left(text_line, 2) = "10" Then ma_security_assets = Mid(text_line, 6)

					If left(text_line, 2) = "11" Then ma_vehicle = Mid(text_line, 6)

					If left(text_line, 2) = "12" Then ma_real_assets = Mid(text_line, 6)

					If left(text_line, 2) = "13" Then
						If left(text_line, 8) = "13 - DET" Then
							other_changes_reported = Mid(text_line, 12)
						Else
							ma_other_changes = Mid(text_line, 6)
						End If

					End If
					If left(text_line, 2) = "15" Then
						If left(text_line, 9) = "15 - RENT" Then
							new_rent_or_mortgage_amount = Mid(text_line, 13)
						ElseIf left(text_line, 9) = "15 - HEAT" Then
							heat_ac_checkbox = Mid(text_line, 13)
						ElseIf left(text_line, 9) = "15 - ELEC" Then
							electricity_checkbox = Mid(text_line, 13)
						ElseIf left(text_line, 9) = "15 - TELE" Then
							telephone_checkbox = Mid(text_line, 13)
						Else
							quest_fifteen_form_answer = Mid(text_line, 6)
						End If
					End If
					If left(text_line, 2) = "16" Then quest_sixteen_form_answer = Mid(text_line, 6)

					If left(text_line, 2) = "17" Then quest_seventeen_form_answer = Mid(text_line, 6)

					If left(text_line, 2) = "18" Then quest_eighteen_form_answer = Mid(text_line, 6)

					If left(text_line, 2) = "19" Then quest_nineteen_form_answer = Mid(text_line, 6)

					'Since the information for each array ittem was joined together we can split it and then use the location in the temporary array toput it in the right place in the main array
					If left(text_line, 17) = "ALL_CLIENTS_ARRAY" Then
						array_info = Mid(text_line, 21)
						array_info = split(array_info, "~")
						ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, known_memb)
						ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) = array_info(0)
						ALL_CLIENTS_ARRAY(memb_last_name, known_memb) = array_info(1)
						ALL_CLIENTS_ARRAY(memb_first_name, known_memb) = array_info(2)
						ALL_CLIENTS_ARRAY(memb_age, known_memb) = array_info(3)
						ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = array_info(4)
						ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = array_info(5)
						ALL_CLIENTS_ARRAY(clt_grh_status, known_memb) = array_info(6)
						ALL_CLIENTS_ARRAY(clt_hc_status, known_memb) = array_info(7)
						ALL_CLIENTS_ARRAY(clt_snap_status, known_memb) = array_info(8)
						known_memb = known_memb + 1
					End If
					If left(text_line, 17) = "NEW_MEMBERS_ARRAY" Then
						array_info = Mid(text_line, 21)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_MEMBERS_ARRAY(new_memb_notes, new_hh_memb)
						NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) = array_info(0)
						NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) = array_info(1)
						NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb) = array_info(2)
						NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) = array_info(3)
						NEW_MEMBERS_ARRAY(new_dob, new_hh_memb) = array_info(4)
						NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb) = array_info(5)
						NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) = array_info(6)
						NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) = array_info(7)
						NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) = array_info(8)
						NEW_MEMBERS_ARRAY(new_memb_moved_in, new_hh_memb) = array_info(9)
						NEW_MEMBERS_ARRAY(new_memb_moved_out, new_hh_memb) = array_info(10)
						NEW_MEMBERS_ARRAY(new_memb_notes, new_hh_memb) = array_info(11)
						new_hh_memb = new_hh_memb + 1

					End If
					If left(text_line, 20) = "NEW_MA_REQUEST_ARRAY" Then
						array_info = Mid(text_line, 24)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_MA_REQUEST_ARRAY(ma_request_notes, each_new_memb)
						NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) = array_info(0)
						each_new_memb = each_new_memb + 1
					End If
					If left(text_line, 16) = "NEW_EARNED_ARRAY" Then
						array_info = Mid(text_line, 20)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_EARNED_ARRAY(earned_notes, each_job)
						NEW_EARNED_ARRAY(earned_client, each_job) = array_info(0)
						NEW_EARNED_ARRAY(earned_type, each_job) = array_info(1)
						NEW_EARNED_ARRAY(earned_source, each_job) = array_info(2)
						NEW_EARNED_ARRAY(earned_change_date, each_job) = array_info(3)
						NEW_EARNED_ARRAY(earned_amount, each_job) = array_info(4)
						NEW_EARNED_ARRAY(earned_freq, each_job) = array_info(5)
						NEW_EARNED_ARRAY(earned_hours, each_job) = array_info(6)
						NEW_EARNED_ARRAY(earned_prog_list, each_job) = array_info(7)
						NEW_EARNED_ARRAY(earned_start_date, each_job) = array_info(8)
						NEW_EARNED_ARRAY(earned_seasonal, each_job) = array_info(9)
						NEW_EARNED_ARRAY(earned_notes, each_job) = array_info(10)
						each_job = each_job + 1
					End If
					If left(text_line, 18) = "NEW_UNEARNED_ARRAY" Then
						array_info = Mid(text_line, 22)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, each_unea)
						NEW_UNEARNED_ARRAY(unearned_client, each_unea) = array_info(0)
						NEW_UNEARNED_ARRAY(unearned_type, each_unea) = array_info(1)
						NEW_UNEARNED_ARRAY(unearned_source, each_unea) = array_info(2)
						NEW_UNEARNED_ARRAY(unearned_change_date, each_unea) = array_info(3)
						NEW_UNEARNED_ARRAY(unearned_amount, each_unea) = array_info(4)
						NEW_UNEARNED_ARRAY(unearned_freq, each_unea) = array_info(5)
						NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = array_info(6)
						NEW_UNEARNED_ARRAY(unearned_start_date, each_unea) = array_info(7)
						NEW_UNEARNED_ARRAY(unearned_notes, each_unea) = array_info(8)
						each_unea = each_unea + 1
					End If
					If left(text_line, 15) = "NEW_ASSET_ARRAY" Then
						array_info = Mid(text_line, 19)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_ASSET_ARRAY(asset_notes, each_asset)
						NEW_ASSET_ARRAY(asset_client, each_asset) = array_info(0)
						NEW_ASSET_ARRAY(asset_type, each_asset) = array_info(1)
						NEW_ASSET_ARRAY(asset_acct_type, each_asset) = array_info(2)
						NEW_ASSET_ARRAY(asset_bank_name, each_asset) = array_info(3)
						NEW_ASSET_ARRAY(asset_year_make_model, each_asset) = array_info(4)
						NEW_ASSET_ARRAY(asset_address, each_asset) = array_info(5)
						NEW_ASSET_ARRAY(asset_prog_list, each_asset) = array_info(6)
						NEW_ASSET_ARRAY(asset_notes, each_asset) = array_info(7)
						each_asset = each_asset + 1
					End If
					If left(text_line, 23) = "NEW_CHILD_SUPPORT_ARRAY" Then
						array_info = Mid(text_line, 27)
						array_info = split(array_info, "~")
						ReDim Preserve NEW_CHILD_SUPPORT_ARRAY(cs_notes, the_cs)
						NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) = array_info(0)
						NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs) = array_info(1)
						NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) = array_info(2)
						NEW_CHILD_SUPPORT_ARRAY(cs_notes, the_cs) = array_info(3)
						the_cs = the_cs + 1
					End If
				Next
			End If
		End If
	End With
end function

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Because the variables are all defined within a function - we need to define them here so that they will be completed outside of the function
Dim MAXIS_footer_month, MAXIS_footer_year, snap_active_count, hc_active_count, grh_active_count, snap_sr_yn, snap_sr_mo, snap_sr_yr, hc_sr_yn, hc_sr_mo, hc_sr_yr, grh_sr_yn, grh_sr_mo, grh_sr_yr, client_on_csr_form
Dim residence_address_match_yn, mailing_address_match_yn, homeless_status, grh_sr, hc_sr, snap_sr, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, new_mail_zip
Dim quest_two_move_in_out, new_hh_memb_not_in_mx_yn, apply_for_ma, q_4_details_blank_checkbox, ma_self_employed, q_5_details_blank_checkbox, ma_start_working, q_6_details_blank_checkbox, ma_other_income
Dim q_7_details_blank_checkbox, ma_liquid_assets, q_9_details_blank_checkbox, ma_security_assets, q_10_details_blank_checkbox, ma_vehicle, q_11_details_blank_checkbox, ma_real_assets, q_12_details_blank_checkbox
Dim ma_other_changes, other_changes_reported, changes_reported_blank_checkbox, quest_fifteen_form_answer, new_rent_or_mortgage_amount, heat_ac_checkbox, electricity_checkbox, telephone_checkbox, shel_proof_provided
Dim quest_sixteen_form_answer, q_16_details_blank_checkbox, quest_seventeen_form_answer, q_17_details_blank_checkbox, quest_eighteen_form_answer, q_18_details_blank_checkbox, quest_nineteen_form_answer, csr_form_date
Dim addr_verif, addr_homeless, addr_reservation, living_situation_status, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, new_mail_one, new_mail_city, new_mail_state
Dim client_signed_yn, client_dated_yn, confirm_csr_form_information, notes_on_faci, notes_on_wreg, new_addr_effective_date, new_resi_one, new_resi_city, new_resi_state, new_resi_zip, new_resi_county, new_shel_verif
Dim new_resi_addr_entered, new_mail_addr_entered, change_resi_addr_btn, change_mail_addr_btn, all_the_clients, full_hh_list

HH_memb_row = 5
Dim row
Dim col

const memb_ref_numb                 = 00
const memb_last_name                = 01
const memb_first_name               = 02
const memb_age                      = 03
const memb_remo_checkbox            = 04
const memb_new_checkbox             = 05
const clt_grh_status                = 06
const clt_hc_status                 = 07
const clt_snap_status               = 08
const memb_id_verif                 = 09
const memb_soc_sec_numb             = 10
const memb_ssn_verif                = 11
const memb_dob                      = 12
const memb_dob_verif                = 13
const memb_gender                   = 14
const memb_rel_to_applct            = 15
const memb_spoken_language          = 16
const memb_written_language         = 17
const memb_interpreter              = 18
const memb_alias                    = 19
const memb_ethnicity                = 20
const memb_race                     = 21
const memi_marriage_status          = 22
const memi_spouse_ref               = 23
const memi_spouse_name              = 24
const memi_designated_spouse        = 25
const memi_marriage_date            = 26
const memi_marriage_verif           = 27
const memi_citizen                  = 28
const memi_citizen_verif            = 29
const memi_last_grade               = 30
const memi_in_MN_less_12_mo         = 31
const memi_resi_verif               = 32
const memi_MN_entry_date            = 33
const memi_former_state             = 34
const memi_other_FS_end             = 35
const wreg_pwe                      = 36
const wreg_status                   = 37
const wreg_defer_fset               = 38
const wreg_fset_orient_date         = 39
const wreg_sanc_begin_date          = 40
const wreg_sanc_count               = 41
const wreg_sanc_reasons             = 42
const wreg_abawd_status             = 43
const wreg_banekd_months            = 44
const wreg_GA_basis                 = 45
const wreg_GA_coop                  = 46
const wreg_numb_ABAWD_months        = 47
const wreg_ABAWD_months_list        = 48
const wreg_numb_second_set_months   = 49
const wreg_second_set_months_list   = 50
Const wreg_notes                    = 51

const shel_hud_sub_yn               = 52
const shel_shared_yn                = 53
const shel_paid_to                  = 54
const shel_rent_retro_amt           = 55
const shel_rent_retro_verif         = 56
const shel_rent_prosp_amt           = 57
const shel_rent_prosp_verif         = 58
const shel_lot_rent_retro_amt       = 59
const shel_lot_rent_retro_verif     = 60
const shel_lot_rent_prosp_amt       = 61
const shel_lot_rent_prosp_verif     = 62
const shel_mortgage_retro_amt       = 63
const shel_mortgage_retro_verif     = 64
const shel_mortgage_prosp_amt       = 65
const shel_mortgage_prosp_verif     = 66
const shel_insurance_retro_amt      = 67
const shel_insurance_retro_verif    = 68
const shel_insurance_prosp_amt      = 69
const shel_insurance_prosp_verif    = 70
const shel_tax_retro_amt            = 71
const shel_tax_retro_verif          = 72
const shel_tax_prosp_amt            = 73
const shel_tax_prosp_verif          = 74
const shel_room_retro_amt           = 75
const shel_room_retro_verif         = 76
const shel_room_prosp_amt           = 77
const shel_room_prosp_verif         = 78
const shel_garage_retro_amt         = 79
const shel_garage_retro_verif       = 80
const shel_garage_prosp_amt         = 81
const shel_garage_prosp_verif       = 82
const shel_subsidy_retro_amt        = 83
const shel_subsidy_retro_verif      = 84
const shel_subsidy_prosp_amt        = 85
const shel_subsidy_prosp_verif      = 86
const shel_notes                    = 87
const shel_verif_checkbox           = 88
const shel_verif_added              = 89
const shel_verif_time               = 90

const memb_notes                    = 91

const new_last_name         = 0
const new_first_name        = 1
const new_mid_initial       = 2
const new_suffix            = 3
const new_full_name         = 4
const new_dob               = 5
const new_rel_to_applicant  = 6
const new_ma_request        = 7
const new_fs_request        = 8
const new_grh_request       = 9
const new_memb_moved_in     = 10
const new_memb_moved_out    = 11
const new_memb_notes        = 12

const ma_request_client     = 0
const ma_request_notes      = 10

Dim NEW_MA_REQUEST_ARRAY()
ReDim NEW_MA_REQUEST_ARRAY(ma_request_notes, 0)


const earned_client         = 0
const earned_type           = 1
const earned_source         = 2
const earned_change_date    = 3
const earned_amount         = 4
const earned_freq           = 5
const earned_hours          = 6
const earned_prog_list      = 7
const earned_start_date     = 8
const earned_seasonal       = 9
const earned_notes          = 11

const unearned_client       = 0
const unearned_type         = 1
const unearned_source       = 2
const unearned_change_date  = 3
const unearned_amount       = 4
const unearned_freq         = 5
Const unearned_prog_list    = 6
const unearned_start_date   = 7
const unearned_notes        = 10

const asset_client          = 0
const asset_type            = 1
const asset_acct_type       = 2
const asset_bank_name       = 3
const asset_year_make_model = 4
const asset_address         = 5
const asset_prog_list       = 9
const asset_notes           = 10

const cs_payer              = 0
const cs_amount             = 1
const cs_current            = 2
const cs_notes              = 10

Const end_of_doc = 6

Dim NEW_EARNED_ARRAY
Dim NEW_UNEARNED_ARRAY
Dim NEW_CHILD_SUPPORT_ARRAY
Dim NEW_ASSET_ARRAY
Dim NEW_MEMBERS_ARRAY
Dim ALL_CLIENTS_ARRAY
ReDim NEW_EARNED_ARRAY(earned_notes, 0)
ReDim NEW_UNEARNED_ARRAY(unearned_notes, 0)
ReDim NEW_CHILD_SUPPORT_ARRAY(cs_notes, 0)
ReDim NEW_ASSET_ARRAY(asset_notes, 0)
ReDim NEW_MEMBERS_ARRAY(new_memb_notes, 0)
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)

csr_form_date = date & ""
Call find_user_name(worker_name)						'defaulting the name of the suer running the script
show_buttons_on_confirmation_dlg = TRUE

Call remove_dash_from_droplist(state_list)

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Initial dialog to gather case number and footer month.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 90, "Case number dialog"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 70, 25, 20, 15, MAXIS_footer_month
  EditBox 95, 25, 20, 15, MAXIS_footer_year
  EditBox 70, 45, 115, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 85, 65, 50, 15
    CancelButton 140, 65, 50, 15
  Text 20, 10, 45, 10, "Case number:"
  Text 25, 30, 40, 10, "CSR Month:"
  Text 125, 30, 25, 10, "mm/yy"
  Text 10, 50, 60, 10, "Worker Signature"
EndDialog
'Showing the case number dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

Call back_to_SELF
vars_filled = FALSE
'This is where we restore the variables if the script for this case is restarted.
Call restore_your_work(vars_filled)

'This is defined within the 'restore_your_work' function. If the variables are restored, we do not need to read this Information
'This saves time and ensures that the variables and arrays are not overwritten.
If vars_filled = FALSE Then
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	'Defining the array we will use to look through all the people real quick.
	'This is similar to the functionality in the FuncLib but it does not allow the script user to select the people
	'Since this is the completion of the CSR - we should be looking at ALL the people.
	Dim HH_member_array()
	ReDim HH_member_array(0)

	hh_count = 0					'this is an incrementer
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadScreen access_denied_check, 13, 24, 2
		If access_denied_check <> "ACCESS DENIED" Then
			ReDim Preserve HH_member_array(hh_count)
			HH_member_array(hh_count) = ref_nbr
			hh_count = hh_count + 1
		End If
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	'This is functionality to read which programs are active on this case
	'Currently this is not being used but I will leave this here because it might be useful.
	Call navigate_to_MAXIS_screen("STAT", "PROG")
	EMReadScreen GRH_status, 4, 9, 74
	EMReadScreen SNAP_status, 4, 10, 74
	EMReadScreen HC_status, 4, 12, 74

	GRH_active = FALSE
	SNAP_active = FALSE
	HC_active = FALSE
	show_buttons_on_confirmation_dlg = TRUE

	If GRH_status = "ACTV" Then GRH_active = TRUE
	If SNAP_status = "ACTV" Then SNAP_active = TRUE
	If HC_status = "ACTV" Then HC_active = TRUE

	'We are now reading for SR information by program and sets the dialog 'autofil'
	Call navigate_to_MAXIS_screen("STAT", "REVW")

	grh_sr = FALSE
	snap_sr = FALSE
	hc_sr = FALSE

	grh_sr_mo = ""
	grh_sr_yr = ""
	EMReadScreen grh_revw_status, 1, 7, 40
	grh_sr_yn = "No"
	If grh_revw_status <> "_" Then
	    EMWriteScreen "X", 5, 35
	    transmit
	    EMReadScreen sr_month, 2, 9, 26
	    EMReadScreen sr_year, 2, 9, 32
	    PF3

	    If grh_revw_status = "N" or grh_revw_status = "I" Then
	        grh_sr_mo = sr_month
	        grh_sr_yr = sr_year
	    Else
	        grh_sr_mo = sr_month
	        sr_year = sr_year * 1
	        sr_year = sr_year - 1
	        grh_sr_yr = right("00" & sr_year, 2)
	    End If
	    grh_sr_yn= "Yes"
	End If

	snap_sr_mo = ""
	snap_sr_yr = ""
	curr_snap_sr_status = ""
	EMReadScreen snap_revw_status, 1, 7, 60
	snap_sr_yn = "No"
	If snap_revw_status <> "_" Then
	    EMWriteScreen "X", 5, 58
	    transmit
	    EMReadScreen sr_month, 2, 9, 26
	    EMReadScreen sr_year, 2, 9, 32
	    PF3
	    If snap_revw_status = "N" or snap_revw_status = "I" Then
	        snap_sr_mo = sr_month
	        snap_sr_yr = sr_year
	    Else
	        snap_sr_mo = sr_month
	        sr_year = sr_year * 1
	        sr_year = sr_year - 1
	        snap_sr_yr = right("00" & sr_year, 2)
	    End If
	    snap_sr_yn= "Yes"
	End If

	hc_sr_mo = ""
	hc_sr_yr = ""
	curr_hc_sr_status = ""
	EMReadScreen hc_revw_status, 1, 7, 73
	hc_sr_yn = "No"
	If hc_revw_status <> "_" Then
	    EMWriteScreen "X", 5, 71
	    transmit
	    EMReadScreen ir_month, 2, 8, 27
	    EMReadScreen ir_year, 2, 8, 33
	    EMReadScreen ar_month, 2, 8, 71
	    EMReadScreen ar_year, 2, 8, 77
	    PF3
	    If ir_month <> "__" Then
	        sr_month = ir_month
	        sr_year = ir_year
	    End If
	    If ar_month <> "__" Then
	        sr_month = ar_month
	        sr_year = ar_year
	    End If
	    If hc_revw_status = "N" or hc_revw_status = "I" Then
	        hc_sr_mo = sr_month
	        hc_sr_yr = sr_year
	    Else
	        hc_sr_mo = sr_month
	        sr_year = sr_year * 1
	        sr_year = sr_year - 1
	        hc_sr_yr = right("00" & sr_year, 2)
	    End If
	    hc_sr_yn= "Yes"
	End If

	Call back_to_SELF

	Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)			'Then we create an array of the the full hh list for looping purpoases
	full_hh_list = Split(list_for_array, chr(9))

	Call back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "MEMB")

	'Now we start filling in the full client array for use in the dialogs
	member_counter = 0
	Do
	    EMReadScreen clt_ref_nbr, 2, 4, 33
	    EMReadScreen clt_last_name, 25, 6, 30
	    EMReadScreen clt_first_name, 12, 6, 63
	    EMReadScreen clt_age, 3, 8, 76

	    ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, member_counter)
	    ALL_CLIENTS_ARRAY(memb_ref_numb, member_counter) = clt_ref_nbr
	    ALL_CLIENTS_ARRAY(memb_last_name, member_counter) = replace(clt_last_name, "_", "")
	    ALL_CLIENTS_ARRAY(memb_first_name, member_counter) = replace(clt_first_name, "_", "")
	    ALL_CLIENTS_ARRAY(memb_age, member_counter) = trim(clt_age)

	    member_counter = member_counter + 1
	    transmit
	    EMReadScreen last_memb, 7, 24, 2
	Loop until last_memb = "ENTER A"

	'Now we gather the address information that exists in MAXIS
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, living_situation_status, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, curr_phone_one, curr_phone_two, curr_phone_three, curr_phone_type_one, curr_phone_type_two, curr_phone_type_three)
End If

'These definitions make sure that in creating of the dialogs there is no confusion in which button is pushed.
back_to_dlg_addr		= 1201
back_to_dlg_ma_income	= 1202
back_to_dlg_ma_asset	= 1203
back_to_dlg_snap		= 1204
back_to_dlg_sig			= 1205
add_another_new_memb_btn= 1206
done_adding_new_memb_btn= 1207
add_memb_btn			= 1208
add_jobs_btn			= 1209
add_unea_btn			= 1210
why_answer_btn			= 1211
next_page_ma_btn		= 1212
add_acct_btn			= 1213
add_secu_btn			= 1214
add_cars_btn			= 1215
add_rest_btn			= 1216
back_to_ma_dlg_1		= 1217
continue_btn			= 1218
back_to_ma_dlg_1		= 1219
back_to_ma_dlg_2		= 1220
finish_ma_questions		= 1221
add_snap_earned_income_btn = 1222
add_snap_unearned_btn	= 1223
add_snap_cs_btn			= 1224
complete_csr_questions	= 1225

next_page_ma_btn = 1100
previous_page_btn = 1200
continue_btn = 1300

'Setting some starting definitions so we can loop around between dialogs
show_csr_dlg_q_1 		= TRUE
show_csr_dlg_q_2 		= TRUE
show_csr_dlg_q_4_7 		= TRUE
show_csr_dlg_q_9_12 	= TRUE
show_csr_dlg_q_13 		= TRUE
show_csr_dlg_q_15_19 	= TRUE
show_csr_dlg_sig 		= TRUE
show_confirmation		= TRUE

csr_dlg_q_1_cleared 	= FALSE
csr_dlg_q_2_cleared 	= FALSE
csr_dlg_q_4_7_cleared 	= FALSE
csr_dlg_q_9_12_cleared 	= FALSE
csr_dlg_q_13_cleared 	= FALSE
csr_dlg_q_15_19_cleared = FALSE
csr_dlg_sig_cleared 	= FALSE

first_q_1_round = TURE
first_q_2_round = TRUE
questions_answered = FALSE
details_shown = FALSE
abawd_on_case = FALSE

'Incrementer starting values
new_memb_counter = 0
new_earned_counter = 0
new_unearned_counter = 0
new_asset_counter = 0


'This massive nested looping fun is so we can go between all these dialogs smoothly and easily
'Each dialog has its own function so that this code is easier to read
Do
	Do
		Do
			Do
				Do
					Do
						Do
							Do
								Do
									show_confirmation = TRUE
									If csr_dlg_q_1_cleared = FALSE Then show_csr_dlg_q_1 = TRUE
									If csr_dlg_q_2_cleared = FALSE Then show_csr_dlg_q_2 = TRUE
									If csr_dlg_q_4_7_cleared = FALSE Then show_csr_dlg_q_4_7 = TRUE
									If csr_dlg_q_9_12_cleared = FALSE Then show_csr_dlg_q_9_12 = TRUE
									If csr_dlg_q_13_cleared = FALSE Then show_csr_dlg_q_13 = TRUE
									If csr_dlg_q_15_19_cleared = FALSE Then show_csr_dlg_q_15_19 = TRUE
									If csr_dlg_sig_cleared = FALSE Then show_csr_dlg_sig = TRUE

									If show_csr_dlg_q_1 = TRUE Then Call csr_dlg_q_1

									If first_q_1_round = TURE and vars_filled = FALSE Then
										Call gather_pers_detail
										first_q_1_round = FALSE
									End If
								Loop until show_csr_dlg_q_1 = FALSE
								save_your_work
								If show_csr_dlg_q_2 = TRUE Then Call csr_dlg_q_2

								If first_q_2_round = TURE Then
									Call count_actives
									first_q_2_round = FALSE
								End If
							Loop until show_csr_dlg_q_2 = FALSE
							save_your_work
							If show_csr_dlg_q_4_7 = TRUE Then Call csr_dlg_q_4_7
						Loop until show_csr_dlg_q_4_7 = FALSE
						save_your_work
						If show_csr_dlg_q_9_12 = TRUE Then Call csr_dlg_q_9_12
					Loop until show_csr_dlg_q_9_12 = FALSE
					save_your_work
					If show_csr_dlg_q_13 = TRUE Then Call csr_dlg_q_13
				Loop until show_csr_dlg_q_13 = FALSE
				save_your_work
				If show_csr_dlg_q_15_19 = TRUE Then Call csr_dlg_q_15_19
			Loop until show_csr_dlg_q_15_19 = FALSE
			save_your_work
			If show_csr_dlg_sig = TRUE Then Call csr_dlg_sig
		Loop until show_csr_dlg_sig = FALSE
		save_your_work
		If show_confirmation = TRUE Then Call confirm_csr_form_dlg
	Loop until confirm_csr_form_information = "YES - This is the information on the CSR Form"
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE
save_your_work

'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
objWord.Caption = "CSR Form Details - CASE #" & MAXIS_case_number			'Title of the document
objWord.Visible = True														'Let the worker see the document

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "CSR Information"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information
objSelection.TypeText "Date Completed: " & date & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR

'Program SR information
If snap_sr_yn = "Yes" Then objSelection.TypeText "SNAP SR for " & snap_sr_mo & "/" & snap_sr_yr & vbCr
If hc_sr_yn = "Yes" Then objSelection.TypeText "HC SR for " & hc_sr_mo & "/" & hc_sr_yr & vbCr
If grh_sr_yn = "Yes" Then objSelection.TypeText "GRH SR for " & grh_sr_mo & "/" & grh_sr_yr & vbCr
objSelection.Font.Size = "11"

'Answers entered for ALL the questions
objSelection.TypeText "CSR Questions:" & vbCr
objSelection.TypeText "Q 1. Name and Address" & vbCr
objSelection.TypeText chr(9) & "Name: " & client_on_csr_form & vbCr
If new_resi_addr_entered = TRUE Then objSelection.TypeText chr(9) & "NEW Address: " & new_resi_one & " " & new_resi_city & ", " & new_resi_state & " " & new_resi_zip & vbCr
If new_resi_addr_entered = FALSE Then objSelection.TypeText chr(9) & "Address: " & resi_line_one & " " & resi_line_two & " " & resi_city & ", " & resi_state & " " & resi_zip & vbCr

If new_mail_addr_entered = TRUE Then objSelection.TypeText chr(9) & "NEW Mailing Address: " & new_mail_one & " " & new_mail_city & ", " & new_mail_state & " " & new_mail_zip & vbCr
If new_mail_addr_entered = FALSE Then objSelection.TypeText chr(9) & "Mailing Address: " & mail_line_one & " " & mail_line_two & " " & mail_city & ", " & mail_state & " " & mail_zip & vbCr
If homeless_status = "Yes" Then  objSelection.TypeText "Reports Homeless" & vbCr

objSelection.TypeText "Q 2. Has anyone moved in or out of your home in the past six months?" & vbCr
objSelection.TypeText chr(9) & quest_two_move_in_out & vbCr
'Listing all the clients that have moved in or out
For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
	If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked OR ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then
		objSelection.TypeText chr(9) & "- MEMB " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & ", age: " & ALL_CLIENTS_ARRAY(memb_age, known_memb)
		If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then objSelection.TypeText " - MOVED IN"  & vbCr
		If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then objSelection.TypeText " - MOVED OUT" & vbCr
	End If
Next
For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
	If NEW_MEMBERS_ARRAY(new_memb_moved_in, new_hh_memb) = checked OR NEW_MEMBERS_ARRAY(new_memb_moved_out, new_hh_memb) = checked Then
		objSelection.TypeText chr(9) & "- "  & ALL_CLIENTS_ARRAY(new_first_name, known_memb) & " "
		If NEW_MEMBERS_ARRAY(new_mid_initial, new_memb_counter) <> "" THen objSelection.TypeText NEW_MEMBERS_ARRAY(new_mid_initial, new_memb_counter) & ". "
		objSelection.TypeText ALL_CLIENTS_ARRAY(new_last_name, known_memb)
		If NEW_MEMBERS_ARRAY(new_suffix, new_memb_counter) <> "" Then objSelection.TypeText NEW_MEMBERS_ARRAY(new_suffix, new_memb_counter)
		If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_hh_memb) = checked Then objSelection.TypeText " - MOVED OUT"  & vbCr
		If NEW_MEMBERS_ARRAY(new_memb_moved_in, new_hh_memb) = checked Then objSelection.TypeText " - MOVED IN"  & vbCr
		objSelection.TypeText chr(9) & chr(9) & "DOB: " & NEW_MEMBERS_ARRAY(new_dob, new_memb_counter) & vbCr
		objSelection.TypeText chr(9) & chr(9) & "Relationship to Applicant: " & NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter) & vbCr
		If NEW_MEMBERS_ARRAY(new_ma_request, new_memb_counter) = checked Then objSelection.TypeText chr(9) & chr(9) & "- Requesting MA" & vbCr
		If NEW_MEMBERS_ARRAY(new_fs_request, new_memb_counter) = checked Then objSelection.TypeText chr(9) & chr(9) & "- Requestiong SNAP" & vbCr
		If NEW_MEMBERS_ARRAY(new_grh_request, new_memb_counter) = checked Then objSelection.TypeText chr(9) & chr(9) & "- Requesting GRH" & vbCr
	End If
Next
table_counter = 1		'Each array is definined by an index - we need to change the index dynamically based on which questions need a table entered. This is the incrementer
objSelection.TypeText "Q 4. Do you want to apply for someone who is not getting coverage now?" & vbCr
objSelection.TypeText chr(9) & apply_for_ma & vbCr
q_4_count = 0		'the count lists how many rows will be in the column
For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
	If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
		found_q_4_det = TRUE	'this identifies if we need to add a table
		q_4_count = q_4_count + 1
	End If
Next
If found_q_4_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_4_count + 1, 1		'This sets the rows and columns needed row then column'
	set objQ4Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ4Table.Cell(1, 1).Range.Text = "Members "
	For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
		If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
			objQ4Table.Cell(tbl_row, 1).Range.Text = NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If

objSelection.TypeText "Q 5. Is anyone self-employed or does anyone expect to be self-employed?" & vbCr
objSelection.TypeText chr(9) & ma_self_employed & vbCr
q_5_count = 1
For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
	If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
		found_q_5_det = TRUE
		q_5_count = q_5_count + 1
	End If
Next
If found_q_5_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_5_count, 4			'This sets the rows and columns needed row then column'
	set objQ5Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ5Table.Cell(1, 1).Range.Text = "Name"
	objQ5Table.Cell(1, 2).Range.Text = "Business"
	objQ5Table.Cell(1, 3).Range.Text = "Start Date"
	objQ5Table.Cell(1, 4).Range.Text = "Yearly Income"
	For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
		If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
			objQ5Table.Cell(tbl_row, 1).Range.Text = NEW_EARNED_ARRAY(earned_client, each_busi)
			objQ5Table.Cell(tbl_row, 2).Range.Text = NEW_EARNED_ARRAY(earned_source, each_busi)
			objQ5Table.Cell(tbl_row, 3).Range.Text = NEW_EARNED_ARRAY(earned_start_date, each_busi)
			objQ5Table.Cell(tbl_row, 4).Range.Text =  NEW_EARNED_ARRAY(earned_amount, each_busi)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If

objSelection.TypeText "Q 6. Does anyone work or does anyone expect to start working?" & vbCr
objSelection.TypeText chr(9) & ma_start_working & vbCr
q_6_count = 0
For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
	If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
		found_q_6_det = TRUE
		q_6_count = q_6_count + 1
	End If
Next
If found_q_6_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_6_count + 1, 6		'This sets the rows and columns needed row then column'
	set objQ6Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ6Table.Cell(1, 1).Range.Text = "Name"
	objQ6Table.Cell(1, 2).Range.Text = "Employer Name"
	objQ6Table.Cell(1, 3).Range.Text = "Start Date"
	objQ6Table.Cell(1, 4).Range.Text = "Seasonal"
	objQ6Table.Cell(1, 5).Range.Text = "Amount Recerived"
	objQ6Table.Cell(1, 6).Range.Text = "How often paid?"
	For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
		If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
			objQ6Table.Cell(tbl_row, 1).Range.Text = NEW_EARNED_ARRAY(earned_client, each_job)
			objQ6Table.Cell(tbl_row, 2).Range.Text = NEW_EARNED_ARRAY(earned_source, each_job)
			objQ6Table.Cell(tbl_row, 3).Range.Text = NEW_EARNED_ARRAY(earned_start_date, each_job)
			objQ6Table.Cell(tbl_row, 4).Range.Text = NEW_EARNED_ARRAY(earned_seasonal, each_job)
			objQ6Table.Cell(tbl_row, 5).Range.Text = NEW_EARNED_ARRAY(earned_amount, each_job)
			objQ6Table.Cell(tbl_row, 6).Range.Text = NEW_EARNED_ARRAY(earned_freq, each_job)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 7. Does anyone get money or does anyone expect to get money from sources other than work?" & vbCr
objSelection.TypeText chr(9) & ma_other_income & vbCr
q_7_count = 0
For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
	If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
		found_q_7_det = TRUE
		q_7_count = q_7_count + 1
	End If
Next
If found_q_7_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_7_count + 1, 5		'This sets the rows and columns needed row then column'
	set objQ7Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ7Table.Cell(1, 1).Range.Text = "Name"
	objQ7Table.Cell(1, 2).Range.Text = "Type of Income"
	objQ7Table.Cell(1, 3).Range.Text = "Start Date"
	objQ7Table.Cell(1, 4).Range.Text = "Amount"
	objQ7Table.Cell(1, 5).Range.Text = "How often Recerived"
	For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
			objQ7Table.Cell(tbl_row, 1).Range.Text = NEW_UNEARNED_ARRAY(unearned_client, each_unea)
			objQ7Table.Cell(tbl_row, 2).Range.Text = NEW_UNEARNED_ARRAY(unearned_source, each_unea)
			objQ7Table.Cell(tbl_row, 3).Range.Text = NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)
			objQ7Table.Cell(tbl_row, 4).Range.Text = NEW_UNEARNED_ARRAY(unearned_amount, each_unea)
			objQ7Table.Cell(tbl_row, 5).Range.Text = NEW_UNEARNED_ARRAY(unearned_freq, each_unea)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 9. Does anyone have cash, a savings or checking account, or certificates of deposit?" & vbCr
objSelection.TypeText chr(9) & ma_liquid_assets & vbCr
q_9_count = 0
For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
	If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		found_q_9_det = TRUE
		q_9_count = q_9_count + 1
	End If
Next
If found_q_9_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_9_count + 1, 3		'This sets the rows and columns needed row then column'
	set objQ9Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ9Table.Cell(1, 1).Range.Text = "Owner(s) Name"
	objQ9Table.Cell(1, 2).Range.Text = "Type"
	objQ9Table.Cell(1, 3).Range.Text = "Name of bank"
	For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			objQ9Table.Cell(tbl_row, 1).Range.Text = NEW_ASSET_ARRAY(asset_client, each_asset)
			objQ9Table.Cell(tbl_row, 2).Range.Text = NEW_ASSET_ARRAY(asset_acct_type, each_asset)
			objQ9Table.Cell(tbl_row, 3).Range.Text = NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 10. Does anyone own or co-own stocks, bonds, retirement accounts, life insurance, burial contracts, annuities, trusts, contracts for deed, or other assets?" & vbCr
objSelection.TypeText chr(9) & ma_security_assets & vbCr
q_10_count = 0
For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
	If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		found_q_10_det = TRUE
		q_10_count = q_10_count + 1
	End If
Next
If found_q_10_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_10_count + 1, 3		'This sets the rows and columns needed row then column'
	set objQ10Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ10Table.Cell(1, 1).Range.Text = "Owner(s) Name"
	objQ10Table.Cell(1, 2).Range.Text = "Type of asset"
	objQ10Table.Cell(1, 3).Range.Text = "Name of company or bank"
	For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			objQ10Table.Cell(tbl_row, 1).Range.Text = NEW_ASSET_ARRAY(asset_client, each_asset)
			objQ10Table.Cell(tbl_row, 2).Range.Text = NEW_ASSET_ARRAY(asset_acct_type, each_asset)
			objQ10Table.Cell(tbl_row, 3).Range.Text = NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 11. Does anyone own a vehicle?" & vbCr
objSelection.TypeText chr(9) & ma_vehicle & vbCr
q_11_count = 0
For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
	If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		found_q_11_det = TRUE
		q_11_count = q_11_count + 1
	End If
Next
If found_q_11_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_11_count + 1, 3		'This sets the rows and columns needed row then column'
	set objQ11Table = objDoc.Tables(table_counter)		'Creates the table with the specific index'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ11Table.Cell(1, 1).Range.Text = "Owner(s) Name"
	objQ11Table.Cell(1, 2).Range.Text = "Type of vehicle"
	objQ11Table.Cell(1, 3).Range.Text = "Year/Make/Model"
	For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			objQ11Table.Cell(tbl_row, 1).Range.Text = NEW_ASSET_ARRAY(asset_client, each_asset)
			objQ11Table.Cell(tbl_row, 2).Range.Text = NEW_ASSET_ARRAY(asset_acct_type, each_asset)
			objQ11Table.Cell(tbl_row, 3).Range.Text = NEW_ASSET_ARRAY(asset_year_make_model, each_asset)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 12. Does anyone own or co-own a home, life estate, cabin, land, time share, rental property or any real estate?" & vbCr
objSelection.TypeText chr(9) & ma_real_assets & vbCr
q_12_count = 0
For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
	If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		found_q_12_det = TRUE
		q_12_count = q_12_count + 1
	End If
Next
If found_q_12_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_12_count + 1, 3		'This sets the rows and columns needed row then column'
	set objQ12Table = objDoc.Tables(table_counter)		'This sets the rows and columns needed row then column'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ12Table.Cell(1, 1).Range.Text = "Owner(s) Name"
	objQ12Table.Cell(1, 2).Range.Text = "Address"
	objQ12Table.Cell(1, 3).Range.Text = "Type of property"
	For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			objQ12Table.Cell(tbl_row, 1).Range.Text = NEW_ASSET_ARRAY(asset_client, each_asset)
			objQ12Table.Cell(tbl_row, 2).Range.Text = NEW_ASSET_ARRAY(asset_address, each_asset)
			objQ12Table.Cell(tbl_row, 3).Range.Text = NEW_ASSET_ARRAY(asset_acct_type, each_asset)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 13. Do you have any changes to report?" & vbCr
objSelection.TypeText chr(9) & ma_other_changes & vbCr
objSelection.TypeText chr(9) & "Details: " & other_changes_reported & vbCr
objSelection.TypeText "Q 15. Has your household moved since your last application or in the past six months?" & vbCr
objSelection.TypeText chr(9) & quest_fifteen_form_answer & vbCr
objSelection.TypeText chr(9) & "New Rent or Mortgage Payment: " & new_rent_or_mortgage_amount & vbCr
objSelection.TypeText chr(9) & "Utilities: " & vbCr
objSelection.TypeText chr(9) & chr(9) & "Heat/Air Conditioning: "
If heat_ac_checkbox = checked then objSelection.TypeText "YES" & vbCr
If heat_ac_checkbox = unchecked then objSelection.TypeText "No" & vbCr
objSelection.TypeText chr(9) & chr(9) & "Electricity: "
If electricity_checkbox = checked then objSelection.TypeText "YES" & vbCr
If electricity_checkbox = unchecked then objSelection.TypeText "No" & vbCr
objSelection.TypeText chr(9) & chr(9) & "Telephone: "
If telephone_checkbox = checked then objSelection.TypeText "YES" & vbCr
If telephone_checkbox = unchecked then objSelection.TypeText "No" & vbCr

objSelection.TypeText "Q 16. Since your last application or in the past six months, has anyone had a change in their income from work such as salary or hourly rate of pay, source of income, starting, stopping or changing jobs, employment status from full-time or part-time? " & vbCr
objSelection.TypeText chr(9) & quest_sixteen_form_answer & vbCr
q_16_count = 0
For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
	If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
		found_q_16_det = TRUE
		q_16_count = q_16_count + 1
	End If
Next
If found_q_16_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_16_count + 1, 6		'This sets the rows and columns needed row then column'
	set objQ16Table = objDoc.Tables(table_counter)		'This sets the rows and columns needed row then column'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ16Table.Cell(1, 1).Range.Text = "Name"
	objQ16Table.Cell(1, 2).Range.Text = "Employer or Business name"
	objQ16Table.Cell(1, 3).Range.Text = "Start or end date"
	objQ16Table.Cell(1, 4).Range.Text = "Amount received"
	objQ16Table.Cell(1, 5).Range.Text = "How often paid"
	objQ16Table.Cell(1, 6).Range.Text = "Hours worked"
	For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
		If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
			objQ16Table.Cell(tbl_row, 1).Range.Text = NEW_EARNED_ARRAY(earned_client, the_earned)
			objQ16Table.Cell(tbl_row, 2).Range.Text = NEW_EARNED_ARRAY(earned_source, the_earned)
			objQ16Table.Cell(tbl_row, 3).Range.Text = NEW_EARNED_ARRAY(earned_change_date, the_earned)
			objQ16Table.Cell(tbl_row, 4).Range.Text = NEW_EARNED_ARRAY(earned_amount, the_earned)
			objQ16Table.Cell(tbl_row, 5).Range.Text = NEW_EARNED_ARRAY(earned_freq, the_earned)
			objQ16Table.Cell(tbl_row, 6).Range.Text = NEW_EARNED_ARRAY(earned_hours, the_earned)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 17. Since your last application or in the past six months, has anyone had a change of more than $50 per month from income sources other than work or a change in a source of unearned income?" & vbCr
objSelection.TypeText chr(9) & quest_seventeen_form_answer & vbCr
q_17_count = 0
For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
	If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
		found_q_17_det = TRUE
		q_17_count = q_17_count + 1
	End If
Next
If found_q_17_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_17_count + 1, 5		'This sets the rows and columns needed row then column'
	set objQ17Table = objDoc.Tables(table_counter)		'This sets the rows and columns needed row then column'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ17Table.Cell(1, 1).Range.Text = "Name"
	objQ17Table.Cell(1, 2).Range.Text = "Type and source of income"
	objQ17Table.Cell(1, 3).Range.Text = "Start or end date"
	objQ17Table.Cell(1, 4).Range.Text = "Amount"
	objQ17Table.Cell(1, 5).Range.Text = "How often received"
	For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
			objQ17Table.Cell(tbl_row, 1).Range.Text = NEW_UNEARNED_ARRAY(unearned_client, the_unearned)
			objQ17Table.Cell(tbl_row, 2).Range.Text = NEW_UNEARNED_ARRAY(unearned_source, the_unearned)
			objQ17Table.Cell(tbl_row, 3).Range.Text = NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)
			objQ17Table.Cell(tbl_row, 4).Range.Text = NEW_UNEARNED_ARRAY(unearned_amount, the_unearned)
			objQ17Table.Cell(tbl_row, 5).Range.Text = NEW_UNEARNED_ARRAY(unearned_freq, the_unearned)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 18. Since your last application or in the past six months, has anyone had a change in court-ordered child or medical support payments?" & vbCr
objSelection.TypeText chr(9) & quest_eighteen_form_answer & vbCr
q_18_count = 0
For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
	If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
		found_q_18_det = TRUE
		q_18_count = q_18_count + 1
	End If
Next
If found_q_18_det = TRUE Then
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, q_18_count + 1, 3		'This sets the rows and columns needed row then column'
	set objQ18Table = objDoc.Tables(table_counter)		'This sets the rows and columns needed row then column'
	table_counter = table_counter + 1					'incrementing the index
	tbl_row = 2											'we start at row 2 with the array information because row 1 is a header row
	objQ18Table.Cell(1, 1).Range.Text = "Name of person paying"
	objQ18Table.Cell(1, 2).Range.Text = "Monthly Amount"
	objQ18Table.Cell(1, 3).Range.Text = "Currently Paying?"
	For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
		If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
			objQ18Table.Cell(tbl_row, 1).Range.Text = NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs)
			objQ18Table.Cell(tbl_row, 2).Range.Text = NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)
			objQ18Table.Cell(tbl_row, 3).Range.Text = NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)							'This adds the borders to the table and formats it
	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If
objSelection.TypeText "Q 19. Did you work 20 hours each week, for an average of 80 hours each month during the past six months?" & vbCr
objSelection.TypeText chr(9) & quest_nineteen_form_answer & vbCr

'Final information for the document
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE
objSelection.TypeText "Verbal Signature accepted on " & csr_form_date

'Here we are creating the file path and saving the file
file_safe_date = replace(date, "/", "-")		'dates cannot have / for a file name so we change it to a -
'We set the file path and name based on case number and date. We can add other criteria if important.
'This MUST have the 'pdf' file extension to work
pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CSR Forms for ECF\CSR - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
'Now we save the document.
'MS Word allows us to save directly as a PDF instead of a DOC.
'the file path must be PDF
'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
objDoc.SaveAs pdf_doc_path, 17

'Now we interact with the system again
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'This looks to see if the PDF file has been correctly saved. If it has the file will exists in the pdf file path
If objFSO.FileExists(pdf_doc_path) = TRUE Then
	'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
	objDoc.Close wdDoNotSaveChanges
	objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"
	'this is the file for the 'save your work' functionality.
	local_work_save_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"
	'we are checking the save your work text file. If it exists we need to delete it because we don't want to save that information locally.
	If objFSO.FileExists(local_work_save_path) = True then
		objFSO.DeleteFile(local_work_save_path)			'DELETE
	End If

	'Now we case note!
	Call start_a_blank_case_note
	Call write_variable_in_CASE_NOTE("CSR Form completed via Phone")
	Call write_variable_in_CASE_NOTE("Form information taken verbally per COVID Waiver Allowance.")
	Call write_variable_in_CASE_NOTE("Form information taken on " & csr_form_date)
	Call write_variable_in_CASE_NOTE("CSR for " & MAXIS_footer_month & "/" & MAXIS_footer_year)
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	'setting the end message
	end_msg = "Success! The information you have provided for the CSR form has been saved to the Assignments forlder so the CSR Form can be updated and added to ECF. The case can be processed using the information saved in the PDF. Additional notes and information are needed or case processing. This script has NOT updated MAXIS or added CSR processing notes."

	'Now we ask if the worker would like the PDF to be opened by the script before the script closes
	'This is helpful because they may not be familiar with where these are saved and they could work from the PDF to process the reVw
	reopen_pdf_doc_msg = MsgBox("The information about the CSR has been saved to a PDF on the LAN to be added to the DHS form and added to ECF." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbYesNo, "Open PDF Doc?")
	If reopen_pdf_doc_msg = vbYes Then
		run_path = chr(34) & pdf_doc_path & chr(34)
		wshshell.Run run_path
		end_msg = end_msg & vbCr & vbCr & "The PDF has been opened for you to view the information that has been saved."
	End If
Else
	end_msg = "Something has gone wrong - the CSR information has NOT been saved correctly to be processed." & vbCr & vbCr & "You can either save the Word Document that has opened as a PDF in the Assignment folder OR Close that document without saving and RERUN the script. Your details have been saved and the script can reopen them and attampt to create the files again. When the script is running, it is best to not interrupt the process."
End If

call script_end_procedure(end_msg)
