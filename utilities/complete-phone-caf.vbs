'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - COMPLETE PHONE CAF.vbs"
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

        If type_one = "H" Then type_one = "Home"
        If type_one = "W" Then type_one = "Work"
        If type_one = "C" Then type_one = "Cell"
        If type_one = "M" Then type_one = "Message"
        If type_one = "T" Then type_one = "TTY/TDD"
        If type_one = "_" Then type_one = ""

        If type_two = "H" Then type_two = "Home"
        If type_two = "W" Then type_two = "Work"
        If type_two = "C" Then type_two = "Cell"
        If type_two = "M" Then type_two = "Message"
        If type_two = "T" Then type_two = "TTY/TDD"
        If type_two = "_" Then type_two = ""

        If type_three = "H" Then type_three = "Home"
        If type_three = "W" Then type_three = "Work"
        If type_three = "C" Then type_three = "Cell"
        If type_three = "M" Then type_three = "Message"
        If type_three = "T" Then type_three = "TTY/TDD"
        If type_three = "_" Then type_three = ""
    End If

    If access_type = "WRITE" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")

        PF9

        Call create_mainframe_friendly_date(addr_eff_date, 4, 43, "YY")

        If len(resi_line_one) > 22 Then
            resi_words = split(resi_line_one, " ")
            write_resi_line_one = ""
            write_resi_line_two = ""
            For each word in resi_words
                If write_resi_line_one = "" Then
                    write_resi_line_one = word
                ElseIf len(write_resi_line_one & " " & word) =< 22 Then
                    write_resi_line_one = write_resi_line_one & " " & word
                Else
                    If write_resi_line_two = "" Then
                        write_resi_line_two = word
                    Else
                        write_resi_line_two = write_resi_line_two & " " & word
                    End If
                End If
            Next
        Else
            write_resi_line_one = resi_line_one
        End If
        EMWriteScreen write_resi_line_one, 6, 43
        EMWriteScreen write_resi_line_two, 7, 43
        EMWriteScreen resi_city, 8, 43
        ' resi_county
        EMWriteScreen left(resi_state, 2), 8, 66
        EMWriteScreen resi_zip, 9, 43

        EMWriteScreen left(addr_verif, 2), 9, 66


        If len(mail_line_one) > 22 Then
            mail_words = split(mail_line_one, " ")
            write_mail_line_one = ""
            write_mail_line_two = ""
            For each word in mail_words
                If write_mail_line_one = "" Then
                    write_mail_line_one = word
                ElseIf len(write_mail_line_one & " " & word) =< 22 Then
                    write_mail_line_one = write_mail_line_one & " " & word
                Else
                    If write_mail_line_two = "" Then
                        write_mail_line_two = word
                    Else
                        write_mail_line_two = write_mail_line_two & " " & word
                    End If
                End If
            Next
        Else
            write_mail_line_one = mail_line_one
        End If
        EMWriteScreen write_mail_line_one, 13, 43
        EMWriteScreen write_mail_line_two, 14, 43
        EMWriteScreen mail_city, 15, 43
        If write_mail_line_one <> "" Then EMWriteScreen left(mail_state, 2), 16, 43
        EMWriteScreen mail_zip, 16, 52

        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

        EMWriteScreen phone_one_left, 17, 45
        EMWriteScreen phone_one_mid, 17, 51
        EMWriteScreen phone_one_right, 17, 55
        If type_one <> "Select ..." Then EMWriteScreen type_one, 17, 67

        EMWriteScreen phone_two_left, 18, 45
        EMWriteScreen phone_two_mid, 18, 51
        EMWriteScreen phone_two_right, 18, 55
        If type_two <> "Select ..." Then EMWriteScreen type_two, 18, 67

        EMWriteScreen phone_three_left, 19, 45
        EMWriteScreen phone_three_mid, 19, 51
        EMWriteScreen phone_three_right, 19, 55
        If type_three <> "Select ..." Then EMWriteScreen type_three, 19, 67

        save_attempt = 1
        Do
            transmit
            EMReadScreen resi_standard_note, 33, 24, 2
            If resi_standard_note = "RESIDENCE ADDRESS IS STANDARDIZED" Then transmit
            EMReadScreen mail_standard_note, 31, 24, 2
            If mail_standard_note = "MAILING ADDRESS IS STANDARDIZED" Then transmit

            row = 0
            col = 0
            EMSearch "Warning:", row, col

            If row <> 0 Then
                Do
                    EMReadScreen warning_note, 55, row, col
                    warning_note = trim(warning_note)
                    warning_message = warning_message & "; " & warning_note
                Loop until warning_note = ""
            End If

            save_attempt = save_attempt + 1
        Loop until save_attempt = 20

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 356, 160, "ADDR Updated"
          EditBox 60, 120, 290, 15, notes_on_address
          ButtonGroup ButtonPressed
            OkButton 300, 140, 50, 15
          Text 10, 10, 160, 10, "The ADDR panel has been updated successfully. "
          Text 10, 30, 155, 20, "When saving the information to the panel, the following warning message was displayed:"
          Text 30, 55, 310, 55, warning_message
          Text 5, 125, 50, 10, "Address Notes:"
        EndDialog

        Do
            err_msg = ""
            dialog Dialog1
            cancel_confirmation

            EMReadScreen addr_check, 4, 2, 44
            If addr_check = "ADDR" Then
                EMReadScreen info_saved, 7, 24, 2
                If info_saved <> "ENTER A"  Then err_msg = err_msg & vbNewLine & "* Review the ADDR panel and update as needed. It appears the script is unable to complete the update without assistance. In order to prevent all work from being lost, please complete the ADDR update manually and press 'OK' for the script to continue once the address information has been saved."
            End If

            If err_msg <> "" Then MsgBox "The ADDR Update functionality needs assistance" & vbNewLine & err_msg
        Loop until err_msg = ""
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
	    ' MsgBox the_snap_status & vbNewLine & person_counter
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
		  DropListBox 185, 60, 75, 45, state_list, new_resi_state
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
		  DropListBox 185, 55, 75, 45, state_list, new_mail_state
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

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then
			.DeleteFile(local_changelog_path)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

		If .FileExists(local_changelog_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)

			'Write the contents of the text file
			objTextStream.WriteLine

			'Close the object so it can be opened again shortly
			objTextStream.Close

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

function restore_your_work(vars_filled)

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			If pull_variables = vbYes Then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_csr_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE

				array_counters = 0
				For Each text_line in saved_csr_details

				Next
			End If
		End If

	End With

end function

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

const memb_ref_numb                 = 00
const memb_last_name                = 01
const memb_first_name               = 02
const memb_mid_name					= 03
const memb_other_names				= 04
const memb_age                      = 05
const memb_remo_checkbox            = 06
const memb_new_checkbox             = 07
const clt_grh_status                = 08
const clt_hc_status                 = 09
const clt_snap_status               = 10
const memb_id_verif                 = 11
const memb_soc_sec_numb             = 12
const memb_ssn_verif                = 13
const memb_dob                      = 14
const memb_dob_verif                = 15
const memb_gender                   = 16
const memb_rel_to_applct            = 17
const memb_spoken_language          = 18
const memb_written_language         = 19
const memb_interpreter              = 20
const memb_alias                    = 21
const memb_ethnicity                = 22
const memb_race                     = 23
const memb_race_a_checkbox			= 24
const memb_race_b_checkbox			= 25
const memb_race_n_checkbox			= 26
const memb_race_p_checkbox			= 27
const memb_race_w_checkbox			= 28
const memi_marriage_status          = 29
const memi_spouse_ref               = 30
const memi_spouse_name              = 31
const memi_designated_spouse        = 32
const memi_marriage_date            = 33
const memi_marriage_verif           = 34
const memi_citizen                  = 35
const memi_citizen_verif            = 36
const memi_last_grade               = 37
const memi_in_MN_less_12_mo         = 38
const memi_resi_verif               = 39
const memi_MN_entry_date            = 40
const memi_former_state             = 41
const memi_other_FS_end             = 42
const clt_snap_checkbox				= 43
const clt_cash_checkbox				= 44
const clt_emer_checkbox				= 45
const clt_none_checkbox 			= 46
const clt_nav_btn					= 47
const clt_intend_to_reside_mn		= 48
const clt_imig_status				= 49
const clt_sponsor_yn 				= 50
const clt_verif_yn					= 51
const clt_verif_details				= 52

const memb_notes                    = 91

Const end_of_doc = 6

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)

state_list = "Select One..."
state_list = state_list+chr(9)+"AL Alabama"
state_list = state_list+chr(9)+"AK Alaska"
state_list = state_list+chr(9)+"AZ Arizona"
state_list = state_list+chr(9)+"AR Arkansas"
state_list = state_list+chr(9)+"CA California"
state_list = state_list+chr(9)+"CO Colorado"
state_list = state_list+chr(9)+"CT Connecticut"
state_list = state_list+chr(9)+"DE Delaware"
state_list = state_list+chr(9)+"DC District Of Columbia"
state_list = state_list+chr(9)+"FL Florida"
state_list = state_list+chr(9)+"GA Georgia"
state_list = state_list+chr(9)+"HI Hawaii"
state_list = state_list+chr(9)+"ID Idaho"
state_list = state_list+chr(9)+"IL Illnois"
state_list = state_list+chr(9)+"IN Indiana"
state_list = state_list+chr(9)+"IA Iowa"
state_list = state_list+chr(9)+"KS Kansas"
state_list = state_list+chr(9)+"KY Kentucky"
state_list = state_list+chr(9)+"LA Louisiana"
state_list = state_list+chr(9)+"ME Maine"
state_list = state_list+chr(9)+"MD Maryland"
state_list = state_list+chr(9)+"MA Massachusetts"
state_list = state_list+chr(9)+"MI Michigan"
state_list = state_list+chr(9)+"MN Minnesota"
state_list = state_list+chr(9)+"MS Mississippi"
state_list = state_list+chr(9)+"MO Missouri"
state_list = state_list+chr(9)+"MT Montana"
state_list = state_list+chr(9)+"NE Nebraska"
state_list = state_list+chr(9)+"NV Nevada"
state_list = state_list+chr(9)+"NH New Hampshire"
state_list = state_list+chr(9)+"NJ New Jersey"
state_list = state_list+chr(9)+"NM New Mexico"
state_list = state_list+chr(9)+"NY New York"
state_list = state_list+chr(9)+"NC North Carolina"
state_list = state_list+chr(9)+"ND North Dakota"
state_list = state_list+chr(9)+"OH Ohio"
state_list = state_list+chr(9)+"OK Oklahoma"
state_list = state_list+chr(9)+"OR Oregon"
state_list = state_list+chr(9)+"PA Pennsylvania"
state_list = state_list+chr(9)+"RI Rhode Island"
state_list = state_list+chr(9)+"SC South Carolina"
state_list = state_list+chr(9)+"SD South Dakota"
state_list = state_list+chr(9)+"TN Tennessee"
state_list = state_list+chr(9)+"TX Texas"
state_list = state_list+chr(9)+"UT Utah"
state_list = state_list+chr(9)+"VT Vermont"
state_list = state_list+chr(9)+"VA Virginia"
state_list = state_list+chr(9)+"WA Washington"
state_list = state_list+chr(9)+"WV West Virginia"
state_list = state_list+chr(9)+"WI Wisconsin"
state_list = state_list+chr(9)+"WY Wyoming"
state_list = state_list+chr(9)+"PR Puerto Rico"
state_list = state_list+chr(9)+"VI Virgin Islands"

memb_panel_relationship_list = "Select One..."
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"01 Applicant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"02 Spouse"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"03 Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"04 Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"05 Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"06 Step Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"08 Step Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"09 Step Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"10 Aunt"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"11 Uncle"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"12 Niece"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"13 Nephew"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"14 Cousin"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"15 Grandparent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"16 Grandchild"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"17 Other Relative"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"18 Legal Guardian"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"24 Not Related"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"25 Live-In Attendant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"27 Unknown"

marital_status = "Select One..."
marital_status = marital_status+chr(9)+"N  Never Married"
marital_status = marital_status+chr(9)+"M  Married Living With Spouse"
marital_status = marital_status+chr(9)+"S  Married Living Apart (Sep)"
marital_status = marital_status+chr(9)+"L  Legally Sep"
marital_status = marital_status+chr(9)+"D  Divorced"
marital_status = marital_status+chr(9)+"W  Widowed"

question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Not Required"

Dim who_are_we_completing_the_form_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, exp_pay_heat_checkbox, exp_pay_ac_checkbox, exp_pay_electricity_checkbox, exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_pne_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county

Dim question_1_yn, question_1_notes, question_1_verif_yn, question_1_verif_details
Dim question_2_yn, question_2_notes, question_2_verif_yn, question_2_verif_details
Dim question_3_yn, question_3_notes, question_3_verif_yn, question_3_verif_details
Dim question_4_yn, question_4_notes, question_4_verif_yn, question_4_verif_details
Dim question_5_yn, question_5_notes, question_5_verif_yn, question_5_verif_details
Dim question_6_yn, question_6_notes, question_6_verif_yn, question_6_verif_details
Dim question_7_yn, question_7_notes, question_7_verif_yn, question_7_verif_details
Dim question_8_yn, question_8_notes, question_8_verif_yn, question_8_verif_details
Dim question_9_yn, question_9_notes, question_9_verif_yn, question_9_verif_details
Dim question_10_yn, question_10_notes, question_10_verif_yn, question_10_verif_details
Dim question_11_yn, question_11_notes, question_11_verif_yn, question_11_verif_details
Dim question_12_yn, question_12_notes, question_12_verif_yn, question_12_verif_details
Dim question_13_yn, question_13_notes, question_13_verif_yn, question_13_verif_details
Dim question_14_yn, question_14_notes, question_14_verif_yn, question_14_verif_details
Dim question_15_yn, question_15_notes, question_15_verif_yn, question_15_verif_details
Dim question_16_yn, question_16_notes, question_16_verif_yn, question_16_verif_details
Dim question_17_yn, question_17_notes, question_17_verif_yn, question_17_verif_details
Dim question_18_yn, question_18_notes, question_18_verif_yn, question_18_verif_details
Dim question_19_yn, question_19_notes, question_19_verif_yn, question_19_verif_details
Dim question_20_yn, question_20_notes, question_20_verif_yn, question_20_verif_details
Dim question_21_yn, question_21_notes, question_21_verif_yn, question_21_verif_details
Dim question_22_yn, question_22_notes, question_22_verif_yn, question_22_verif_details
Dim question_23_yn, question_23_notes, question_23_verif_yn, question_23_verif_details
Dim question_24_yn, question_24_notes, question_24_verif_yn, question_24_verif_details
'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
application_date = date & ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 141, 165, "Case Number Information"
  EditBox 85, 40, 50, 15, MAXIS_case_number
  CheckBox 15, 60, 115, 10, "Check here if this is a new case ", no_case_number_checkbox
  EditBox 85, 85, 50, 15, application_date
  EditBox 15, 120, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 30, 145, 50, 15
    CancelButton 85, 145, 50, 15
  Text 10, 10, 125, 25, "This script will guide you through all of the CAF questions to complete the form with the client over the phone."
  Text 35, 45, 50, 10, "Case Number:"
  Text 25, 70, 115, 10, "and there is no Case Number yet."
  Text 25, 90, 60, 10, "Application Date:"
  Text 15, 110, 70, 10, "Sign your case note"
EndDialog
'Showing the case number dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
		If no_case_number_checkbox = checked Then err_msg = ""
        ' Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If IsDate(application_date) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

show_known_addr = FALSE
vars_filled = FALSE

Call back_to_SELF
Call restore_your_work(vars_filled)
Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)
If vars_filled = TRUE Then show_known_addr = TRUE

If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then

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

	Call back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	    transmit

	    EMReadScreen clt_id_verif, 2, 9, 68
	    EMReadScreen clt_ssn, 11, 7, 42
	    EMReadScreen clt_ssn_verif, 1, 7, 68
	    EMReadScreen clt_dob, 10, 8, 42
	    EMReadScreen clt_dob_verif, 2, 8, 68
	    EMReadScreen clt_gender, 1, 9, 42

	    EMReadScreen clt_rel_to_applct, 2, 10, 42
	    EMReadScreen clt_spkn_lang, 20, 12, 42
	    EMReadScreen clt_wrt_lang, 29, 13, 42
	    EMReadScreen clt_interp_need, 1, 14, 68
	    EMReadScreen clt_alias, 1, 15, 42
	    EMReadScreen clt_ethncty, 1, 16, 68
	    EMReadScreen clt_race_sum, 30, 17, 42
		PF9
		EMReadScreen in_edit_mode, 9, 24, 11
		If in_edit_mode <> "READ ONLY" Then
			EMWriteScreen "X", 17, 34
			transmit
			EMReadScreen race_x_a, 1, 7, 12
			EMReadScreen race_x_b, 1, 8, 12
			EMReadScreen race_x_n, 1, 10, 12
			EMReadScreen race_x_p, 1, 12, 12
			EMReadScreen race_x_w, 1, 14, 12
			EMReadScreen race_x_u, 1, 15, 12
			PF10
			PF10
		End If

	    If clt_id_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "BC - Birth Certificate"
	    If clt_id_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "RE - Religious Record"
	    If clt_id_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DL - Drivers Lic/St ID"
	    If clt_id_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DV - Divorce Decree"
	    If clt_id_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "AL - Alien Card"
	    If clt_id_verif = "AD" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "AD - Arrival/Depart"
	    If clt_id_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DR - Doctor Stmt"
	    If clt_id_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "PV = Passport/Visa"
	    If clt_id_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "OT - Other"
	    If clt_id_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "NO - No Ver Prvd"
	    ALL_CLIENTS_ARRAY(memb_soc_sec_numb, case_memb) = replace(clt_ssn, " ", "-")
	    If clt_ssn_verif = "A" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "A - SSN Applied For"
	    If clt_ssn_verif = "P" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "P - SSN Prvd, Verif Pending"
	    If clt_ssn_verif = "N" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "N - SSN Not Prvd"
	    If clt_ssn_verif = "V" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "V - System Verified"
	    ALL_CLIENTS_ARRAY(memb_dob, case_memb) = replace(clt_dob, " ", "/")
	    If clt_dob_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "BC - Birth Certificate"
	    If clt_dob_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "RE - Religious Record"
	    If clt_dob_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DL - Drivers Lic/St ID"
	    If clt_dob_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DV - Divorce Decree"
	    If clt_dob_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "AL - Alien Card"
	    If clt_dob_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DR - Doctor Stmt"
	    If clt_dob_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "PV = Passport/Visa"
	    If clt_dob_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "OT - Other"
	    If clt_dob_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "NO - No Ver Prvd"
	    If clt_gender = "F" Then ALL_CLIENTS_ARRAY(memb_gender, case_memb) = "Female"
	    If clt_gender = "M" Then ALL_CLIENTS_ARRAY(memb_gender, case_memb) = "Male"
	    If clt_rel_to_applct = "01" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "01 Applicant"
	    If clt_rel_to_applct = "02" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "02 Spouse"
	    If clt_rel_to_applct = "03" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "03 Child"
	    If clt_rel_to_applct = "04" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "04 Parent"
	    If clt_rel_to_applct = "05" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "05 Sibling"
	    If clt_rel_to_applct = "06" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "06 Step Sibling"
	    If clt_rel_to_applct = "08" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "08 Step Child"
	    If clt_rel_to_applct = "09" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "09 Step Parent"
	    If clt_rel_to_applct = "10" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "10 Aunt"
	    If clt_rel_to_applct = "11" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "11 Uncle"
	    If clt_rel_to_applct = "12" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "12 Niece"
	    If clt_rel_to_applct = "13" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "13 Nephew"
	    If clt_rel_to_applct = "14" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "14 Cousin"
	    If clt_rel_to_applct = "15" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "15 Grandparent"
	    If clt_rel_to_applct = "16" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "16 Grandchild"
	    If clt_rel_to_applct = "17" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "17 Other Relative"
	    If clt_rel_to_applct = "18" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "18 Legal Guardian"
	    If clt_rel_to_applct = "24" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "24 Not Related"
	    If clt_rel_to_applct = "25" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "25 Live-In Attendant"
	    If clt_rel_to_applct = "27" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "27 Unknown"

	    clt_spkn_lang = replace(clt_spkn_lang, "_", "")
	    clt_spkn_lang = replace(clt_spkn_lang, "  ", " - ")
	    ALL_CLIENTS_ARRAY(memb_spoken_language, case_memb) = trim(clt_spkn_lang)
	    clt_wrt_lang = replace(clt_wrt_lang, "_", "")
	    clt_wrt_lang = replace(clt_wrt_lang, "  ", " - ")
	    clt_wrt_lang = replace(clt_wrt_lang, "(HRF)", "")
	    ALL_CLIENTS_ARRAY(memb_written_language, case_memb) = trim(clt_wrt_lang)

	    ALL_CLIENTS_ARRAY(memb_interpreter, case_memb) = clt_interp_need
	    ALL_CLIENTS_ARRAY(memb_alias, case_memb) = clt_alias
	    ALL_CLIENTS_ARRAY(memb_ethnicity, case_memb) = clt_ethncty
	    ALL_CLIENTS_ARRAY(memb_race, case_memb) = trim(clt_race_sum)

		If race_x_a = "X" Then race_a_checkbox = checked
		If race_x_b = "X" Then race_b_checkbox = checked
		If race_x_n = "X" Then race_n_checkbox = checked
		If race_x_p = "X" Then race_p_checkbox = checked
		If race_x_w = "X" Then race_w_checkbox = checked
		' If race_x_u = "X" Then race_a_checkbox = checked
	Next

	Call navigate_to_MAXIS_screen("STAT", "MEMI")
	For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	    transmit

	    EMReadScreen clt_mar_status, 1, 7, 40
	    EMReadScreen clt_spouse, 2, 9, 49

	    EMReadScreen clt_desg_spouse_yn, 1, 7, 71
	    EMReadScreen clt_marriage_date, 8, 8, 40
	    EMReadScreen clt_marriage_date_verif, 8, 8, 71

	    EMReadScreen clt_citizen, 1, 11, 49
	    EMReadScreen clt_cit_verif, 2, 11, 78
	    EMReadScreen clt_last_grade, 2, 10, 49
	    EMReadScreen clt_in_MN_12_mo, 1, 14, 49
	    EMReadScreen clt_resi_verif, 1, 14, 78
	    EMReadScreen clt_MN_entry_date, 8, 15, 49
	    EMReadScreen clt_former_state, 2, 15, 78
	    EMReadScreen clt_other_st_FS_end, 8, 13, 49

	    If clt_mar_status = "N" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "N Never married"
	    If clt_mar_status = "M" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "M Married, Living with Spouse"
	    If clt_mar_status = "S" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "S Married Living Apart"
	    If clt_mar_status = "L" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "L Legally Separated"
	    If clt_mar_status = "D" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "D Divorced"
	    If clt_mar_status = "W" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "W Widowed"
	    ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) = replace(clt_spouse, "_", "")
	    If ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) <> "" Then
	        For all_the_people = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
	            If ALL_CLIENTS_ARRAY(memb_ref_nbr, all_the_people) = ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) Then
	                ALL_CLIENTS_ARRAY(memi_spouse_name, case_memb) = ALL_CLIENTS_ARRAY(memb_first_name, all_the_people) & " " & ALL_CLIENTS_ARRAY(memb_last_name, all_the_people)
	            End If
	        Next
	    End If
	    ALL_CLIENTS_ARRAY(memi_designated_spouse, case_memb) = replace(clt_desg_spouse_yn, "_", "")
	    ALL_CLIENTS_ARRAY(memi_marriage_date, case_memb) = replace(clt_marriage_date, " ", "/")
	    ALL_CLIENTS_ARRAY(memi_marriage_verif, case_memb) = replace(clt_marriage_date_verif, " ", "/")
	    ALL_CLIENTS_ARRAY(memi_citizen, case_memb) = clt_citizen
	    If clt_cit_verif = "BC" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "BC - Birth Certificate"
	    If clt_cit_verif = "RE" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "RE - Religious Record"
	    If clt_cit_verif = "NP" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "NP - Naturalization Papers"
	    If clt_cit_verif = "IM" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "IM - Immigration Document"
	    If clt_cit_verif = "PV" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "PV - Passport/Visa"
	    If clt_cit_verif = "OT" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "OT - Other Document"
	    If clt_cit_verif = "NO" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "NO - No Ver prvd"

	    If clt_last_grade = "00" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Pre 1st Grd"
	    If clt_last_grade = "01" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 1"
	    If clt_last_grade = "02" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 2"
	    If clt_last_grade = "03" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 3"
	    If clt_last_grade = "04" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 4"
	    If clt_last_grade = "05" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 5"
	    If clt_last_grade = "06" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 6"
	    If clt_last_grade = "07" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 7"
	    If clt_last_grade = "08" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 8"
	    If clt_last_grade = "09" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 9"
	    If clt_last_grade = "10" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 10"
	    If clt_last_grade = "11" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 11"
	    If clt_last_grade = "12" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "HS Diploma or GED"
	    If clt_last_grade = "13" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Some Post Sec Ed"
	    If clt_last_grade = "14" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "High Schl Plus Cert"
	    If clt_last_grade = "15" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Four Yr Degree"
	    If clt_last_grade = "16" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grad Degree"

	    ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = clt_in_MN_12_mo
	    If ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = ""
	    IF clt_resi_verif = "1" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "1 - Rent Receipt"
	    IF clt_resi_verif = "2" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "2 - Landlord's Stmt"
	    IF clt_resi_verif = "3" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "3 - Utility Bill"
	    IF clt_resi_verif = "4" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "4 - Other"
	    IF clt_resi_verif = "N" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "N - Ver Not Prvd"
	    ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = replace(clt_MN_entry_date, " ", "/")
	    If ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = ""
	    ALL_CLIENTS_ARRAY(memi_former_state, case_memb) = replace(clt_former_state, "_", "")
	    ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = replace(clt_other_st_FS_end, " ", "/")
	    If ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = ""

	Next

	'Now we gather the address information that exists in MAXIS
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, mail_line_one, mail_line_two, mail_addr_city, mail_addr_state, mail_addr_zip, addr_eff_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_pne_type, phone_two_type, phone_four_type)
	resi_addr_street_full = resi_line_one & " " & resi_line_two
	resi_addr_street_full = trim(resi_addr_street_full)
	mail_addr_street_full = mail_line_one & " " & mail_line_two
	mail_addr_street_full = trim(mail_addr_street_full)

	show_known_addr = TRUE
End If







function dlg_page_one_pers_and_exp()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 416, 240, "CAF Person and Expedited"
			  DropListBox 205, 10, 205, 45, all_the_clients, who_are_we_completing_the_form_with
			  DropListBox 205, 30, 205, 45, all_the_clients, caf_person_one
			  EditBox 290, 65, 50, 15, exp_q_1_income_this_month
			  EditBox 310, 85, 50, 15, exp_q_2_assets_this_month
			  EditBox 250, 105, 50, 15, exp_q_3_rent_this_month
			  CheckBox 125, 125, 30, 10, "Heat", exp_pay_heat_checkbox
			  CheckBox 160, 125, 65, 10, "Air Conditioning", exp_pay_ac_checkbox
			  CheckBox 230, 125, 45, 10, "Electricity", exp_pay_electricity_checkbox
			  CheckBox 280, 125, 35, 10, "Phone", exp_pay_phone_checkbox
			  CheckBox 325, 125, 35, 10, "None", exp_pay_none_checkbox
			  DropListBox 245, 140, 40, 45, "No"+chr(9)+"Yes", exp_migrant_seasonal_formworker_yn
			  DropListBox 365, 155, 40, 45, "No"+chr(9)+"Yes", exp_received_previous_assistance_yn
			  EditBox 80, 175, 80, 15, exp_previous_assistance_when
			  EditBox 200, 175, 85, 15, exp_previous_assistance_where
			  EditBox 320, 175, 85, 15, exp_previous_assistance_what
			  DropListBox 160, 195, 40, 45, "No"+chr(9)+"Yes", exp_pregnant_yn
			  ComboBox 255, 195, 150, 45, all_the_clients, exp_pregnant_who
			  ButtonGroup ButtonPressed
				PushButton 305, 220, 50, 15, "Next", next_btn
				' PushButton 250, 225, 50, 10, "Back", back_btn
			    CancelButton 360, 220, 50, 15
			  Text 70, 15, 130, 10, "Who are you completing the form with?"
			  Text 145, 35, 55, 10, "Select Person 1:"
			  GroupBox 10, 50, 400, 165, "Expedited Questions - Do you need help right away?"
			  Text 20, 70, 270, 10, "1. How much income (cash or checkes) did or will your household get this month?"
			  Text 20, 90, 290, 10, "2. How much does your household (including children) have cash, checking or savings?"
			  Text 20, 110, 225, 10, "3. How much does your household pay for rent/mortgage per month?"
			  Text 30, 125, 90, 10, "What utilities do you pay?"
			  Text 20, 145, 225, 10, "4. Is anyone in your household a migrant or seasonal farm worker?"
			  Text 20, 160, 345, 10, "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
			  Text 30, 180, 50, 10, "If yes, When?"
			  Text 170, 180, 30, 10, "Where?"
			  Text 295, 180, 25, 10, "What?"
			  Text 20, 200, 135, 10, "6. Is anyone in your household pregnant?"
			  Text 210, 200, 40, 10, "If yes, who?"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

		Loop until ButtonPressed = next_btn
	Loop until err_msg = ""
	If exp_pregnant_who = "Select or Type" Then exp_pregnant_who = ""

	show_caf_pg_1_pers_dlg = FALSE
	caf_pg_1_pers_dlg_cleared = TRUE
end function


function dlg_page_one_address()

	If resi_addr_street_full = blank Then show_known_addr = FALSE
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""

			If show_known_addr = TRUE Then
				BeginDialog Dialog1, 0, 0, 391, 285, "CAF Address"
				  Text 70, 55, 305, 15, resi_addr_street_full
				  Text 70, 75, 105, 15, resi_addr_city
				  Text 205, 75, 110, 45, resi_addr_state
				  Text 340, 75, 35, 15, resi_addr_zip
				  Text 125, 95, 45, 45, reservation_yn
				  Text 245, 85, 130, 15, reservation_name
				  Text 125, 115, 45, 45, homeless_yn
				  Text 245, 115, 130, 45, living_situation
				  Text 70, 155, 305, 15, mail_addr_street_full
				  Text 70, 175, 105, 15, mail_addr_city
				  Text 205, 175, 110, 45, mail_addr_state
				  Text 340, 175, 35, 15, mail_addr_zip
				  Text 20, 225, 90, 15, phone_one_number
				  Text 125, 225, 65, 45, phone_pne_type
				  Text 20, 245, 90, 15, phone_two_number
				  Text 125, 245, 65, 45, phone_two_type
				  Text 20, 265, 90, 15, phone_three_number
				  Text 125, 265, 65, 45, phone_three_type
				  Text 325, 205, 50, 15, address_change_date
				  Text 255, 240, 120, 45, resi_addr_county
				  ButtonGroup ButtonPressed
					PushButton 280, 265, 50, 15, "Next", next_btn
					PushButton 280, 253, 50, 10, "Back", back_btn
				    CancelButton 335, 265, 50, 15
				    PushButton 290, 20, 95, 15, "Update Information", update_information_btn
					' PushButton 290, 20, 95, 15, "Save Information", save_information_btn
				    PushButton 325, 135, 50, 10, "CLEAR", clear_mail_addr_btn
					PushButton 205, 220, 35, 10, "CLEAR", clear_phone_one_btn
				    PushButton 205, 240, 35, 10, "CLEAR", clear_phone_two_btn
				    PushButton 205, 260, 35, 10, "CLEAR", clear_phone_three_btn
				  Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
				  GroupBox 10, 190, 235, 90, "Phone Number"
				  Text 20, 55, 45, 10, "House/Street"
				  Text 45, 75, 20, 10, "City"
				  Text 185, 75, 20, 10, "State"
				  Text 325, 75, 15, 10, "Zip"
				  Text 20, 95, 100, 10, "Do you live on a Reservation?"
				  Text 180, 95, 60, 10, "If yes, which one?"
				  Text 30, 115, 90, 10, "Client Indicates Homeless:"
				  Text 185, 115, 60, 10, "Living Situation?"
				  GroupBox 10, 35, 375, 95, "Residence Address"
				  Text 20, 155, 45, 10, "House/Street"
				  Text 45, 175, 20, 10, "City"
				  Text 185, 175, 20, 10, "State"
				  Text 325, 175, 15, 10, "Zip"
				  GroupBox 10, 125, 375, 70, "Mailing Address"
				  Text 20, 205, 50, 10, "Number"
				  Text 125, 205, 25, 10, "Type"
				  Text 255, 205, 60, 10, "Date of Change:"
				  Text 255, 225, 75, 10, "County of Residence:"
				EndDialog
			End If

			If show_known_addr = FALSE Then
				BeginDialog Dialog1, 0, 0, 391, 285, "CAF Address"
				  EditBox 70, 50, 305, 15, resi_addr_street_full
				  EditBox 70, 70, 105, 15, resi_addr_city
				  DropListBox 205, 70, 110, 45, state_list, resi_addr_state
				  EditBox 340, 70, 35, 15, resi_addr_zip
				  DropListBox 125, 90, 45, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", reservation_yn
				  EditBox 245, 90, 130, 15, reservation_name
				  DropListBox 125, 110, 45, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", homeless_yn
				  DropListBox 245, 110, 130, 45, "Select"+chr(9)+"", living_situation
				  EditBox 70, 150, 305, 15, mail_addr_street_full
				  EditBox 70, 170, 105, 15, mail_addr_city
				  DropListBox 205, 170, 110, 45, state_list, mail_addr_state
				  EditBox 340, 170, 35, 15, mail_addr_zip
				  EditBox 20, 220, 90, 15, phone_one_number
				  DropListBox 125, 220, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_pne_type
				  EditBox 20, 240, 90, 15, phone_two_number
				  DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_two_type
				  EditBox 20, 260, 90, 15, phone_three_number
				  DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", phone_three_type
				  EditBox 325, 200, 50, 15, address_change_date
				  DropListBox 255, 235, 120, 45, county_list, resi_addr_county
				  ButtonGroup ButtonPressed
					PushButton 280, 265, 50, 15, "Next", next_btn
					PushButton 280, 253, 50, 10, "Back", back_btn
				    CancelButton 335, 265, 50, 15
				    ' PushButton 290, 20, 95, 15, "Update Information", update_information_btn
					PushButton 290, 20, 95, 15, "Save Information", save_information_btn
				    PushButton 325, 135, 50, 10, "CLEAR", clear_mail_addr_btn
				    PushButton 205, 220, 35, 10, "CLEAR", clear_phone_one_btn
				    PushButton 205, 240, 35, 10, "CLEAR", clear_phone_two_btn
				    PushButton 205, 260, 35, 10, "CLEAR", clear_phone_three_btn
				  Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
				  GroupBox 10, 190, 235, 90, "Phone Number"
				  Text 20, 55, 45, 10, "House/Street"
				  Text 45, 75, 20, 10, "City"
				  Text 185, 75, 20, 10, "State"
				  Text 325, 75, 15, 10, "Zip"
				  Text 20, 95, 100, 10, "Do you live on a Reservation?"
				  Text 180, 95, 60, 10, "If yes, which one?"
				  Text 30, 115, 90, 10, "Client Indicates Homeless:"
				  Text 185, 115, 60, 10, "Living Situation?"
				  GroupBox 10, 35, 375, 95, "Residence Address"
				  Text 20, 155, 45, 10, "House/Street"
				  Text 45, 175, 20, 10, "City"
				  Text 185, 175, 20, 10, "State"
				  Text 325, 175, 15, 10, "Zip"
				  GroupBox 10, 125, 375, 70, "Mailing Address"
				  Text 20, 205, 50, 10, "Number"
				  Text 125, 205, 25, 10, "Type"
				  Text 255, 205, 60, 10, "Date of Change:"
				  Text 255, 225, 75, 10, "County of Residence:"
				EndDialog
			End If

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = update_information_btn Then show_known_addr = FALSE
			If ButtonPressed = save_information_btn Then show_known_addr = TRUE
			If ButtonPressed = clear_mail_addr_btn Then
				mail_addr_street_full = ""
				mail_addr_city = ""
				mail_addr_state = "Select One..."
				mail_addr_zip = ""
			End If
			If ButtonPressed = clear_phone_one_btn Then
				phone_one_number = ""
				phone_pne_type = "Select One..."
			End If
			If ButtonPressed = clear_phone_two_btn Then
				phone_two_number = ""
				phone_two_type = "Select One..."
			End If
			If ButtonPressed = clear_phone_three_btn Then
				phone_three_number = ""
				phone_three_type = "Select One..."
			End If
			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_1_pers_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_1_addr_dlg = FALSE
		caf_pg_1_addr_dlg_cleared = TRUE
	End If
end function

function dlg_page_two_household_comp()

	known_membs = 0
	shown_known_pers_detail = TRUE
	If ALL_CLIENTS_ARRAY(memb_last_name, known_membs) = "" Then shown_known_pers_detail = FALSE
	go_back = FALSE
	Do
		Do
			btn_placeholder = 3001
			For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb) = btn_placeholder
				btn_placeholder = btn_placeholder + 1
			Next

			err_msg = ""
			Dialog1 = ""

			' If no_case_number_checkbox = checked Then
			'
			' Else
			'
			' End If

			If shown_known_pers_detail = TRUE Then
				BeginDialog Dialog1, 0, 0, 541, 310, "Household Member Information"
				  Text 20, 45, 105, 15, ALL_CLIENTS_ARRAY(memb_last_name, known_membs)
				  Text 130, 45, 90, 15, ALL_CLIENTS_ARRAY(memb_first_name, known_membs)
				  Text 225, 45, 70, 15, ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)
				  Text 300, 45, 175, 15, ALL_CLIENTS_ARRAY(memb_other_names, known_membs)
				  If ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs) = "V - System Verified" Then
					  Text 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  Else
					  EditBox 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  End If
				  Text 95, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_dob, known_membs)
				  Text 170, 75, 50, 45, ALL_CLIENTS_ARRAY(memb_gender, known_membs)
				  Text 225, 75, 140, 45, ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)
				  Text 370, 75, 105, 45, ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)
				  Text 20, 105, 130, 15, ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)
				  Text 155, 105, 70, 15, ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)
				  Text 230, 105, 165, 15, ALL_CLIENTS_ARRAY(memi_former_state, known_membs)
				  Text 400, 105, 75, 45, ALL_CLIENTS_ARRAY(memi_citizen, known_membs)
				  Text 20, 135, 60, 45, ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)
				  Text 90, 135, 170, 15, ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)
				  Text 90, 165, 170, 15, ALL_CLIENTS_ARRAY(memb_written_language, known_membs)
				  Text 280, 155, 40, 45, ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)
				  CheckBox 330, 155, 30, 10, "Asian", ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)
				  CheckBox 330, 165, 30, 10, "Black", ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)
				  CheckBox 330, 175, 120, 10, "American Indian or Alaska Native", ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)
				  CheckBox 330, 185, 130, 10, "Pacific Islander and Native Hawaiian", ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)
				  CheckBox 330, 195, 130, 10, "White", ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)
				  CheckBox 20, 195, 50, 10, "SNAP (food)", ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)
				  CheckBox 75, 195, 65, 10, "Cash programs", ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)
				  CheckBox 145, 195, 85, 10, "Emergency Assistance", ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)
				  CheckBox 235, 195, 30, 10, "NONE", ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)
				  DropListBox 15, 230, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)
				  EditBox 100, 230, 205, 15, ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)
				  DropListBox 310, 230, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)
				  DropListBox 15, 260, 80, 50, "None Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
				  EditBox 100, 260, 435, 15, ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)
				  EditBox 15, 290, 350, 15, ALL_CLIENTS_ARRAY(memb_notes, known_membs)
				  ButtonGroup ButtonPressed
					PushButton 430, 290, 50, 15, "Next", next_btn
					PushButton 375, 295, 50, 10, "Back", back_btn
					CancelButton 485, 290, 50, 15
					PushButton 330, 5, 95, 15, "Update Information", update_information_btn
					' PushButton 330, 5, 95, 15, "Save Information", save_information_btn
					y_pos = 35
					For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
						PushButton 490, y_pos, 45, 10, "Person " & (the_memb + 1), ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb)
						y_pos = y_pos + 10
					Next
					PushButton 490, 230, 45, 10, "Add Person", add_person_btn
				  Text 10, 10, 315, 10, "Known Member in MAXIS - current information. If there are changes, press this button to update:"
				  If ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs) = "" Then
					  GroupBox 10, 25, 475, 190, "Person " & known_membs+1
				  Else
					  GroupBox 10, 25, 475, 190, "Person " & known_membs+1 & " - MEMBER " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				  End If
				  Text 20, 35, 50, 10, "Last Name"
				  Text 130, 35, 50, 10, "First Name"
				  Text 225, 35, 50, 10, "Middle Name"
				  Text 300, 35, 50, 10, "Other Names"
				  Text 20, 65, 55, 10, "Soc Sec Number"
				  Text 95, 65, 45, 10, "Date of Birth"
				  Text 170, 65, 45, 10, "Gender"
				  Text 225, 65, 90, 10, "Relationship to MEMB 01"
				  Text 370, 65, 50, 10, "Marital Status"
				  Text 20, 95, 75, 10, "Last Grade Completed"
				  Text 155, 95, 55, 10, "Moved to MN on"
				  Text 230, 95, 65, 10, "Moved to MN from"
				  Text 400, 95, 75, 10, "US Citizen or National"
				  Text 20, 125, 40, 10, "Interpreter?"
				  Text 90, 125, 95, 10, "Preferred Spoken Language"
				  Text 90, 155, 95, 10, "Preferred Written Language"
				  GroupBox 270, 135, 205, 75, "Demographics - OPTIONAL"
				  Text 280, 145, 35, 10, "Hispanic?"
				  Text 330, 145, 50, 10, "Race"
				  Text 20, 185, 145, 10, "Which programs is this person requesting?"
				  Text 15, 220, 80, 10, "Intends to reside in MN"
				  Text 100, 220, 65, 10, "Immigration Status"
				  Text 310, 220, 50, 10, "Sponsor?"
				  Text 15, 250, 50, 10, "Verification"
				  Text 100, 250, 65, 10, "Verification Details"
				  Text 15, 280, 50, 10, "Notes:"
				EndDialog

			End If

			If shown_known_pers_detail = FALSE Then

				BeginDialog Dialog1, 0, 0, 541, 310, "Household Member Information"
				  EditBox 20, 45, 105, 15, ALL_CLIENTS_ARRAY(memb_last_name, known_membs)
				  EditBox 130, 45, 90, 15, ALL_CLIENTS_ARRAY(memb_first_name, known_membs)
				  EditBox 225, 45, 70, 15, ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)
				  EditBox 300, 45, 175, 15, ALL_CLIENTS_ARRAY(memb_other_names, known_membs)
				  EditBox 20, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)
				  EditBox 95, 75, 70, 15, ALL_CLIENTS_ARRAY(memb_dob, known_membs)
				  DropListBox 170, 75, 50, 45, ""+chr(9)+"Male"+chr(9)+"Female", ALL_CLIENTS_ARRAY(memb_gender, known_membs)
				  DropListBox 225, 75, 140, 45, memb_panel_relationship_list, ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)
				  DropListBox 370, 75, 105, 45, marital_status, ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)
				  EditBox 20, 105, 130, 15, ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)
				  EditBox 155, 105, 70, 15, ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)
				  EditBox 230, 105, 165, 15, ALL_CLIENTS_ARRAY(memi_former_state, known_membs)
				  DropListBox 400, 105, 75, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memi_citizen, known_membs)
				  DropListBox 20, 135, 60, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)
				  EditBox 90, 135, 170, 15, ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)
				  EditBox 90, 165, 170, 15, ALL_CLIENTS_ARRAY(memb_written_language, known_membs)
				  DropListBox 280, 155, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)
				  CheckBox 330, 155, 30, 10, "Asian", ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)
				  CheckBox 330, 165, 30, 10, "Black", ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)
				  CheckBox 330, 175, 120, 10, "American Indian or Alaska Native", ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)
				  CheckBox 330, 185, 130, 10, "Pacific Islander and Native Hawaiian", ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)
				  CheckBox 330, 195, 130, 10, "White", ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)
				  CheckBox 20, 195, 50, 10, "SNAP (food)", ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)
				  CheckBox 75, 195, 65, 10, "Cash programs", ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)
				  CheckBox 145, 195, 85, 10, "Emergency Assistance", ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)
				  CheckBox 235, 195, 30, 10, "NONE", ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)
				  DropListBox 15, 230, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)
				  EditBox 100, 230, 205, 15, ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)
				  DropListBox 310, 230, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)
				  DropListBox 15, 260, 80, 50, "None Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
				  EditBox 100, 260, 435, 15, ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)
				  EditBox 15, 290, 350, 15, ALL_CLIENTS_ARRAY(memb_notes, known_membs)
				  ButtonGroup ButtonPressed
				    PushButton 430, 290, 50, 15, "Next", next_btn
					PushButton 375, 295, 50, 10, "Back", back_btn
				    CancelButton 485, 290, 50, 15
				    ' PushButton 330, 5, 95, 15, "Update Information", update_information_btn
					PushButton 330, 5, 95, 15, "Save Information", save_information_btn
					y_pos = 35
					For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
						PushButton 490, y_pos, 45, 10, "Person " & (the_memb + 1), ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb)
						y_pos = y_pos + 10
					Next
				    PushButton 490, 230, 45, 10, "Add Person", add_person_btn
				  Text 10, 10, 315, 10, "Known Member in MAXIS - current information. If there are changes, press this button to update:"
				  GroupBox 10, 25, 475, 190, "MEMBER " &  ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				  Text 20, 35, 50, 10, "Last Name"
				  Text 130, 35, 50, 10, "First Name"
				  Text 225, 35, 50, 10, "Middle Name"
				  Text 300, 35, 50, 10, "Other Names"
				  Text 20, 65, 55, 10, "Soc Sec Number"
				  Text 95, 65, 45, 10, "Date of Birth"
				  Text 170, 65, 45, 10, "Gender"
				  Text 225, 65, 90, 10, "Relationship to MEMB 01"
				  Text 370, 65, 50, 10, "Marital Status"
				  Text 20, 95, 75, 10, "Last Grade Completed"
				  Text 155, 95, 55, 10, "Moved to MN on"
				  Text 230, 95, 65, 10, "Moved to MN from"
				  Text 400, 95, 75, 10, "US Citizen or National"
				  Text 20, 125, 40, 10, "Interpreter?"
				  Text 90, 125, 95, 10, "Preferred Spoken Language"
				  Text 90, 155, 95, 10, "Preferred Written Language"
				  GroupBox 270, 135, 205, 75, "Demographics - OPTIONAL"
				  Text 280, 145, 35, 10, "Hispanic?"
				  Text 330, 145, 50, 10, "Race"
				  Text 20, 185, 145, 10, "Which programs is this person requesting?"
				  Text 15, 220, 80, 10, "Intends to reside in MN"
				  Text 100, 220, 65, 10, "Immigration Status"
				  Text 310, 220, 50, 10, "Sponsor?"
				  Text 15, 250, 50, 10, "Verification"
				  Text 100, 250, 65, 10, "Verification Details"
				  Text 15, 280, 50, 10, "Notes:"
				EndDialog

			End If

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = next_btn Then
				known_membs = known_membs + 1
				If known_membs =< UBound(ALL_CLIENTS_ARRAY, 2) Then ButtonPressed = ""
			End If
			If ButtonPressed = update_information_btn Then shown_known_pers_detail = FALSE
			If ButtonPressed = save_information_btn Then shown_known_pers_detail = TRUE
			For the_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				If ButtonPressed = ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb) Then known_membs = the_memb
			Next
			If ButtonPressed = back_btn Then
				If known_membs = 0 Then
					go_back = TRUE
					ButtonPressed = next_btn
					err_msg = ""
					show_caf_pg_1_addr_dlg = TRUE
				Else
					known_membs = known_membs - 1
				End If
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_2_hhcomp_dlg = FALSE
		caf_pg_2_hhcomp_dlg_cleared = TRUE
	End If

end function

function dlg_page_three_household_info()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 471, 305, "Tell Us About Your Household"
			  DropListBox 10, 10, 60, 45, question_answers, question_1_yn
			  EditBox 120, 20, 235, 15, question_1_notes
			  DropListBox 10, 45, 60, 45, question_answers, question_2_yn
			  EditBox 120, 65, 235, 15, question_2_notes
			  DropListBox 10, 90, 60, 45, question_answers, question_3_yn
			  EditBox 120, 100, 235, 15, question_3_notes
			  DropListBox 10, 125, 60, 45, question_answers, question_4_yn
			  EditBox 120, 145, 235, 15, question_4_notes
			  DropListBox 10, 170, 60, 45, question_answers, question_5_yn
			  EditBox 120, 190, 235, 15, question_5_notes
			  DropListBox 10, 215, 60, 45, question_answers, question_6_yn
			  EditBox 120, 225, 235, 15, question_6_notes
			  DropListBox 10, 250, 60, 45, question_answers, question_7_yn
			  EditBox 120, 280, 235, 15, question_7_notes
			  ButtonGroup ButtonPressed
			    PushButton 360, 285, 50, 15, "Next", next_btn
			    PushButton 360, 275, 50, 10, "Back", back_btn
			    CancelButton 415, 285, 50, 15
			    PushButton 380, 20, 75, 10, "ADD VERIFICATION", add_verif_1_btn
			    PushButton 380, 55, 75, 10, "ADD VERIFICATION", add_verif_2_btn
			    PushButton 380, 100, 75, 10, "ADD VERIFICATION", add_verif_3_btn
			    PushButton 380, 135, 75, 10, "ADD VERIFICATION", add_verif_4_btn
			    PushButton 380, 180, 75, 10, "ADD VERIFICATION", add_verif_5_btn
			    PushButton 380, 225, 75, 10, "ADD VERIFICATION", add_verif_6_btn
			    PushButton 380, 260, 75, 10, "ADD VERIFICATION", add_verif_7_btn
			  Text 80, 10, 230, 10, "1. Does everyone in your household buy, fix or eat food with you?"
			  Text 95, 25, 25, 10, "Notes:"
			  Text 360, 10, 100, 10, "Q1 - Verification - " & question_1_verif_yn
			  Text 80, 45, 245, 10, "2. Is anyone in the household, who is age 60 or over or disabled, unable to "
			  Text 90, 55, 115, 10, "buy or fix food due to a disability?"
			  Text 95, 70, 25, 10, "Notes:"
			  Text 360, 45, 100, 10, "Q2 - Verification - " & question_2_verif_yn
			  Text 80, 90, 165, 10, "3. Is anyone in the household attending school?"
			  Text 95, 105, 25, 10, "Notes:"
			  Text 360, 90, 100, 10, "Q3 - Verification - " & question_3_verif_yn
			  Text 80, 125, 230, 10, "4. Is anyone in your household temporarily not living in your home?"
			  Text 90, 135, 230, 10, "(for example: vacation, foster care, treatment, hospital, job search)"
			  Text 95, 150, 25, 10, "Notes:"
			  Text 360, 125, 100, 10, "Q4 - Verification - " & question_4_verif_yn
			  Text 80, 170, 255, 10, "5. Is anyone blind, or does anyone have a physical or mental health condition"
			  Text 90, 180, 185, 10, " that limits the ability to work or perform daily activities?"
			  Text 95, 195, 25, 10, "Notes:"
			  Text 360, 170, 100, 10, "Q5 - Verification - " & question_5_verif_yn
			  Text 80, 215, 245, 10, "6. Is anyone unable to work for reasons other than illness or disability?"
			  Text 95, 230, 25, 10, "Notes:"
			  Text 360, 215, 100, 10, "Q6 - Verification - " & question_6_verif_yn
			  Text 80, 250, 170, 10, "7. In the last 60 days did anyone in the household: "
			  Text 90, 260, 165, 20, "- Stop working or quit a job?   - Refuse a job offer? - Ask to work fewer hours?   - Go on strike?"
			  Text 95, 285, 25, 10, "Notes:"
			  Text 360, 250, 100, 10, "Q7 - Verification - " & question_7_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_1_btn Then Call verif_details_dlg(1)
			If ButtonPressed = add_verif_2_btn Then Call verif_details_dlg(2)
			If ButtonPressed = add_verif_3_btn Then Call verif_details_dlg(3)
			If ButtonPressed = add_verif_4_btn Then Call verif_details_dlg(4)
			If ButtonPressed = add_verif_5_btn Then Call verif_details_dlg(5)
			If ButtonPressed = add_verif_6_btn Then Call verif_details_dlg(6)
			If ButtonPressed = add_verif_7_btn Then Call verif_details_dlg(7)

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_2_hhcomp_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_3_hhinfo_dlg = FALSE
		caf_pg_3_hhinfo_dlg_cleared = TRUE
	End If

end function

function verif_details_dlg(question_number)

	' If question_number = 1 Then
	' If question_number = 2 Then
	' If question_number = 3 Then
	' If question_number = 4 Then
	' If question_number = 5 Then
	' If question_number = 6 Then
	' If question_number = 7 Then
	' If question_number = 8 Then
	' If question_number = 9 Then
	' If question_number = 10 Then
	' If question_number = 11 Then
	' If question_number = 12 Then
	' If question_number = 13 Then
	' If question_number = 14 Then
	' If question_number = 15 Then
	' If question_number = 16 Then
	' If question_number = 17 Then
	' If question_number = 18 Then
	' If question_number = 19 Then
	' If question_number = 20 Then
	' If question_number = 21 Then
	' If question_number = 22 Then
	' If question_number = 23 Then
	' If question_number = 24 Then


	Select Case question_number
		Case 1
			verif_selection = question_1_verif_yn
			verif_detials = question_1_verif_details
			question_words = "1. Does everyone in your household buy, fix or eat food with you?"
		Case 2
			verif_selection = question_2_verif_yn
			verif_detials = question_2_verif_details
			question_words = "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		Case 3
			verif_selection = question_3_verif_yn
			verif_detials = question_3_verif_details
			question_words = "3. Is anyone in the household attending school?"
		Case 4
			verif_selection = question_4_verif_yn
			verif_detials = question_4_verif_details
			question_words = "4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)"
		Case 5
			verif_selection = question_5_verif_yn
			verif_detials = question_5_verif_details
			question_words = "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
		Case 6
			verif_selection = question_6_verif_yn
			verif_detials = question_6_verif_details
			question_words = "6. Is anyone unable to work for reasons other than illness or disability?"
		Case 7
			verif_selection = question_7_verif_yn
			verif_detials = question_7_verif_details
			question_words = "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
		Case 8
			verif_selection = question_8_verif_yn
			verif_detials = question_8_verif_details
			question_words = "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
		Case 9
			verif_selection = question_9_verif_yn
			verif_detials = question_9_verif_details
			question_words = "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
		Case 10
			verif_selection = question_10_verif_yn
			verif_detials = question_10_verif_details
			question_words = "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
		Case 11
			verif_selection = question_11_verif_yn
			verif_detials = question_11_verif_details
			question_words = "11. Do you expect any changes in income, expenses or work hours?"
		Case 12
			verif_selection = question_12_verif_yn
			verif_detials = question_12_verif_details
			question_words = "12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
		Case 13
			verif_selection = question_13_verif_yn
			verif_detials = question_13_verif_details
			question_words = "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
		Case 14
			verif_selection = question_14_verif_yn
			verif_detials = question_14_verif_details
			question_words = "14. Does your household have the following housing expenses?"
		Case 15
			verif_selection = question_15_verif_yn
			verif_detials = question_15_verif_details
			question_words = "15. Does your household have the following utility expenses any time during the year?"
		Case 16
			verif_selection = question_16_verif_yn
			verif_detials = question_16_verif_details
			question_words = "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
		Case 17
			verif_selection = question_17_verif_yn
			verif_detials = question_17_verif_details
			question_words = "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
		Case 18
			verif_selection = question_18_verif_yn
			verif_detials = question_18_verif_details
			question_words = "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
		Case 19
			verif_selection = question_19_verif_yn
			verif_detials = question_19_verif_details
			question_words = "19. For SNAP only: Does anyone in the household have medical expenses? "
		Case 20
			verif_selection = question_20_verif_yn
			verif_detials = question_20_verif_details
			question_words = "20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
		Case 21
			verif_selection = question_21_verif_yn
			verif_detials = question_21_verif_details
			question_words = "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)"
		Case 22
			verif_selection = question_22_verif_yn
			verif_detials = question_22_verif_details
			question_words = "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
		Case 23
			verif_selection = question_23_verif_yn
			verif_detials = question_23_verif_details
			question_words = "23. For children under the age of 19, are both parents living in the home?"
		Case 24
			verif_selection = question_24_verif_yn
			verif_detials = question_24_verif_details
			question_words = "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
	End Select


	BeginDialog Dialog1, 0, 0, 396, 95, "Add Verification"
	  DropListBox 60, 35, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", verif_selection
	  EditBox 60, 55, 330, 15, verif_detials
	  ButtonGroup ButtonPressed
	    PushButton 340, 75, 50, 15, "Return", return_btn
		PushButton 145, 35, 50, 10, "CLEAR", clear_btn
	  Text 10, 10, 380, 20, question_words
	  Text 10, 40, 45, 10, "Verification: "
	  Text 20, 60, 30, 10, "Details:"
	EndDialog

	Do
		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = clear_btn Then
			verif_selection = "Not Needed"
			verif_detials = ""
		End If
	Loop until ButtonPressed = return_btn

	Select Case question_number
		Case 1
			question_1_verif_yn = verif_selection
			question_1_verif_details = verif_detials
		Case 2
			question_2_verif_yn = verif_selection
			question_2_verif_details = verif_detials
		Case 3
			question_3_verif_yn = verif_selection
			question_3_verif_details = verif_detials
		Case 4
			question_4_verif_yn = verif_selection
			question_4_verif_details = verif_detials
		Case 5
			question_5_verif_yn = verif_selection
			question_5_verif_details = verif_detials
		Case 6
			question_6_verif_yn = verif_selection
			question_6_verif_details = verif_detials
		Case 7
			question_7_verif_yn = verif_selection
			question_7_verif_details = verif_detials
		Case 8
			question_8_verif_yn = verif_selection
			question_8_verif_details = verif_detials
		Case 9
			question_9_verif_yn = verif_selection
			question_9_verif_details = verif_detials
		Case 10
			question_10_verif_yn = verif_selection
			question_10_verif_details = verif_detials
		Case 11
			question_11_verif_yn = verif_selection
			question_11_verif_details = verif_detials
		Case 12
			question_12_verif_yn = verif_selection
			question_12_verif_details = verif_detials
		Case 13
			question_13_verif_yn = verif_selection
			question_13_verif_details = verif_detials
		Case 14
			question_14_verif_yn = verif_selection
			question_14_verif_details = verif_detials
		Case 15
			question_15_verif_yn = verif_selection
			question_15_verif_details = verif_detials
		Case 16
			question_16_verif_yn = verif_selection
			question_16_verif_details = verif_detials
		Case 17
			question_17_verif_yn = verif_selection
			question_17_verif_details = verif_detials
		Case 18
			question_18_verif_yn = verif_selection
			question_18_verif_details = verif_detials
		Case 19
			question_19_verif_yn = verif_selection
			question_19_verif_details = verif_detials
		Case 20
			question_20_verif_yn = verif_selection
			question_20_verif_details = verif_detials
		Case 21
			question_21_verif_yn = verif_selection
			question_21_verif_details = verif_detials
		Case 22
			question_22_verif_yn = verif_selection
			question_22_verif_details = verif_detials
		Case 23
			question_23_verif_yn = verif_selection
			question_23_verif_details = verif_detials
		Case 24
			question_24_verif_yn = verif_selection
			question_24_verif_details = verif_detials
	End Select

end function

BeginDialog Dialog1, 0, 0, 541, 310, "Household Member Information"
  DropListBox 15, 260, 80, 50, "", verif_yn
  EditBox 100, 260, 435, 15, verif_details
  EditBox 15, 290, 350, 15, notes
  ButtonGroup ButtonPressed
    PushButton 430, 290, 50, 15, "Next", next_btn
    PushButton 375, 295, 50, 10, "Back", back_btn
    CancelButton 485, 290, 50, 15
  Text 15, 250, 50, 10, "Verification"
  Text 100, 250, 65, 10, "Verification Details"
  Text 15, 280, 50, 10, "Notes:"
EndDialog

BeginDialog Dialog1, 0, 0, 391, 285, "Dialog"
  ButtonGroup ButtonPressed
    PushButton 280, 265, 50, 15, "Next", next_btn
    CancelButton 335, 265, 50, 15
    PushButton 290, 20, 95, 15, "Update Information", update_information_btn
  Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
EndDialog

next_btn					= 1000
back_btn					= 1010
update_information_btn		= 1020
save_information_btn		= 1030
clear_mail_addr_btn			= 1040
clear_phone_one_btn			= 1041
clear_phone_two_btn			= 1042
clear_phone_three_btn		= 1043
add_person_btn				= 1050
add_verif_1_btn				= 1060
add_verif_2_btn				= 1061
add_verif_3_btn				= 1062
add_verif_4_btn				= 1063
add_verif_5_btn				= 1064
add_verif_6_btn				= 1065
add_verif_7_btn				= 1066

show_caf_pg_1_pers_dlg = TRUE
show_caf_pg_1_addr_dlg = TRUE
show_caf_pg_2_hhcomp_dlg = TRUE
show_caf_pg_3_hhinfo_dlg = TRUE

caf_pg_1_pers_dlg_cleared = FALSE
caf_pg_1_addr_dlg_cleared = FALSE
caf_pg_2_hhcomp_dlg_cleared = FALSE
caf_pg_3_hhinfo_dlg_cleared = FALSE

Do
	Do
		Do
			Do
				show_confirmation = TRUE
				If caf_pg_1_pers_dlg_cleared = FALSE Then show_caf_pg_1_pers_dlg = TRUE
				If caf_pg_1_addr_dlg_cleared = FALSE Then show_caf_pg_1_addr_dlg = TRUE
				If caf_pg_2_hhcomp_dlg_cleared = FALSE Then show_caf_pg_2_hhcomp_dlg = TRUE
				If caf_pg_3_hhinfo_dlg_cleared = FALSE Then show_caf_pg_3_hhinfo_dlg = TRUE

				If show_caf_pg_1_pers_dlg = TRUE Then Call dlg_page_one_pers_and_exp

			Loop until show_caf_pg_1_pers_dlg = FALSE
			' save_your_work
			If show_caf_pg_1_addr_dlg = TRUE Then Call dlg_page_one_address
		Loop until show_caf_pg_1_addr_dlg = FALSE
		' save_your_work
		If show_caf_pg_2_hhcomp_dlg = TRUE Then Call dlg_page_two_household_comp
	Loop until show_caf_pg_2_hhcomp_dlg = FALSE
	' save_your_work
	If show_caf_pg_3_hhinfo_dlg = TRUE Then Call dlg_page_three_household_info
Loop until show_caf_pg_3_hhinfo_dlg = FALSE

'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "CAF Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "CAF Form Details - CASE #" & MAXIS_case_number			'Title of the document
objWord.Visible = True														'Let the worker see the document

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF Information"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

If MAXIS_case_number <> "" Then objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information
If no_case_number_checkbox = checked Then objSelection.TypeText "New Case - no case number" & vbCr
objSelection.TypeText "Date Completed: " & date & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR

'Program CAF Information
caf_progs = ""
If CAF_requesting_SNAP = "Yes" Then caf_progs = caf_progs & ", SNAP"
If CAF_requesting_CASH = "Yes" Then caf_progs = caf_progs & ", Cash"
If CAF_requesting_EMER = "Yes" Then caf_progs = caf_progs & ", EMER"
If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
objSelection.TypeText "CAF requesting: " & caf_progs & vbCr
objSelection.Font.Size = "11"


'Ennumeration for SetHeight and SetWidth
'wdAdjustFirstColumn	2	Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
	' wdAdjustNone			0	Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
	' wdAdjustProportional	1	Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
	' wdAdjustSameWidth		3	Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.


objSelection.TypeText "PERSON 1"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 16, 1					'This sets the rows and columns needed row then column'
set objPers1Table = objDoc.Tables(1)		'Creates the table with the specific index'

objPers1Table.AutoFormat(16)							'This adds the borders to the table and formats it
objPers1Table.Columns(1).Width = 500

for row = 1 to 15 Step 2
	objPers1Table.Cell(row, 1).SetHeight 10, 2
Next
for row = 2 to 16 Step 2
	objPers1Table.Cell(row, 1).SetHeight 15, 2
Next

For row = 1 to 2
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 140, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 85, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
For col = 1 to 4
	objPers1Table.Cell(1, col).Range.Font.Size = 6
	objPers1Table.Cell(2, col).Range.Font.Size = 12
Next

objPers1Table.Cell(1, 1).Range.Text = "APPLICANT'S LEGAL NAME - LAST"
objPers1Table.Cell(1, 2).Range.Text = "FIRST NAME"
objPers1Table.Cell(1, 3).Range.Text = "MIDDLE NAME"
objPers1Table.Cell(1, 4).Range.Text = "OTHER NAMES YOU USE"

objPers1Table.Cell(2, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_last_name, 0)
objPers1Table.Cell(2, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_first_name, 0)
objPers1Table.Cell(2, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_mid_name, 0)
objPers1Table.Cell(2, 4).Range.Text = ALL_CLIENTS_ARRAY(memb_other_names, 0)

' objPers1Table.Cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 3).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 4).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Borders(wdBorderBottom) = wdLineStyleNone

' objPers1Table.Rows(3).Cells.Split 1, 5, TRUE
For row = 3 to 4
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 110, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 115, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
For col = 1 to 4
	objPers1Table.Cell(3, col).Range.Font.Size = 6
	objPers1Table.Cell(4, col).Range.Font.Size = 12
Next
objPers1Table.Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
objPers1Table.Cell(3, 2).Range.Text = "DATE OF BIRTH"
objPers1Table.Cell(3, 3).Range.Text = "GENDER"
objPers1Table.Cell(3, 4).Range.Text = "MARITAL STATUS"

objPers1Table.Cell(4, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_soc_sec_numb, 0)
objPers1Table.Cell(4, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_dob, 0)
objPers1Table.Cell(4, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_gender, 0)
objPers1Table.Cell(4, 4).Range.Text = ALL_CLIENTS_ARRAY(memi_marriage_status, 0)

' objPers1Table.Rows(4).Cells.Split 1, 5, TRUE
For row = 5 to 6
	objPers1Table.Rows(row).Cells.Split 1, 5, TRUE

	objPers1Table.Cell(row, 1).SetWidth 230, 2
	objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 3).SetWidth 110, 2
	objPers1Table.Cell(row, 4).SetWidth 30, 2
	objPers1Table.Cell(row, 5).SetWidth 75, 2
Next
For col = 1 to 5
	objPers1Table.Cell(5, col).Range.Font.Size = 6
	objPers1Table.Cell(6, col).Range.Font.Size = 12
Next
objPers1Table.Cell(5, 1).Range.Text = "ADDRESS WHERE YOU LIVE"
objPers1Table.Cell(5, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(5, 3).Range.Text = "CITY"
objPers1Table.Cell(5, 4).Range.Text = "STATE"
objPers1Table.Cell(5, 5).Range.Text = "ZIP CODE"

objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full
objPers1Table.Cell(6, 2).Range.Text = ""
objPers1Table.Cell(6, 3).Range.Text = resi_addr_city
objPers1Table.Cell(6, 4).Range.Text = LEFT(resi_addr_state, 2)
objPers1Table.Cell(6, 5).Range.Text = resi_addr_zip


' objPers1Table.Rows(5).Cells.Split 1, 3, TRUE
For row = 7 to 8
	objPers1Table.Rows(row).Cells.Split 1, 5, TRUE

	objPers1Table.Cell(row, 1).SetWidth 230, 2
	objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 3).SetWidth 110, 2
	objPers1Table.Cell(row, 4).SetWidth 30, 2
	objPers1Table.Cell(row, 5).SetWidth 75, 2
Next
For col = 1 to 5
	objPers1Table.Cell(7, col).Range.Font.Size = 6
	objPers1Table.Cell(8, col).Range.Font.Size = 12
Next
objPers1Table.Cell(7, 1).Range.Text = "MAILING ADDRESS"
objPers1Table.Cell(7, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(7, 3).Range.Text = "CITY"
objPers1Table.Cell(7, 4).Range.Text = "STATE"
objPers1Table.Cell(7, 5).Range.Text = "ZIP CODE"

objPers1Table.Cell(8, 1).Range.Text = mail_addr_street_full
objPers1Table.Cell(8, 2).Range.Text = ""
objPers1Table.Cell(8, 3).Range.Text = mail_addr_city
objPers1Table.Cell(8, 4).Range.Text = LEFT(mail_addr_state, 2)
objPers1Table.Cell(8, 5).Range.Text = mail_addr_zip


' objPers1Table.Rows(6).Cells.Split 1, 3, TRUE
For row = 9 to 10
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 115, 2
	objPers1Table.Cell(row, 2).SetWidth 115, 2
	objPers1Table.Cell(row, 3).SetWidth 270, 2
Next
For col = 1 to 3
	objPers1Table.Cell(9, col).Range.Font.Size = 6
	objPers1Table.Cell(10, col).Range.Font.Size = 12
Next
objPers1Table.Cell(9, 1).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 2).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 3).Range.Text = "DO YOU LIVE ON A RESERVATION?"

objPers1Table.Cell(10, 1).Range.Text = phone_one_number & " (" & phone_pne_type & ")"
objPers1Table.Cell(10, 2).Range.Text = phone_two_number & " (" & phone_two_type & ")"
objPers1Table.Cell(10, 3).Range.Text = reservation_yn & " - " & reservation_name


' objPers1Table.Rows(7).Cells.Split 1, 3, TRUE
For row = 11 to 12
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 190, 2
	objPers1Table.Cell(row, 3).SetWidth 190, 2
Next
For col = 1 to 3
	objPers1Table.Cell(11, col).Range.Font.Size = 6
	objPers1Table.Cell(12, col).Range.Font.Size = 12
Next
objPers1Table.Cell(11, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
objPers1Table.Cell(11, 2).Range.Text = "WHAT IS YOU RPREFERRED SPOKEN LANGUAGE?"
objPers1Table.Cell(11, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

objPers1Table.Cell(12, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_interpreter, 0)
objPers1Table.Cell(12, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_spoken_language, 0)
objPers1Table.Cell(12, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_written_language, 0)


' objPers1Table.Rows(8).Cells.Split 1, 3, TRUE
For row = 13 to 14
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 270, 2
	objPers1Table.Cell(row, 3).SetWidth 110, 2
Next
For col = 1 to 3
	objPers1Table.Cell(13, col).Range.Font.Size = 6
	objPers1Table.Cell(14, col).Range.Font.Size = 12
Next
objPers1Table.Cell(13, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
objPers1Table.Cell(13, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
objPers1Table.Cell(13, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

objPers1Table.Cell(14, 1).Range.Text = ALL_CLIENTS_ARRAY(memi_last_grade, 0)
objPers1Table.Cell(14, 2).Range.Text = "Date: " & ALL_CLIENTS_ARRAY(memi_MN_entry_date, 0) & "   From: " & ALL_CLIENTS_ARRAY(memi_former_state, 0)
objPers1Table.Cell(14, 3).Range.Text = ALL_CLIENTS_ARRAY(memi_citizen, 0)

For row = 15 to 16
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 275, 2
	objPers1Table.Cell(row, 2).SetWidth 95, 2
	objPers1Table.Cell(row, 3).SetWidth 130, 2
Next
For col = 1 to 3
	objPers1Table.Cell(15, col).Range.Font.Size = 6
	objPers1Table.Cell(16, col).Range.Font.Size = 12
Next
objPers1Table.Cell(15, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
objPers1Table.Cell(15, 2).Range.Text = "ETHNICITY"
objPers1Table.Cell(15, 3).Range.Text = "RACE"

If ALL_CLIENTS_ARRAY(clt_none_checkbox, 0) = checked then progs_applying_for = "NONE"
If ALL_CLIENTS_ARRAY(clt_snap_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", SNAP"
If ALL_CLIENTS_ARRAY(clt_cash_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Cash"
If ALL_CLIENTS_ARRAY(clt_emer_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

If ALL_CLIENTS_ARRAY(memb_race_a_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Asian"
If ALL_CLIENTS_ARRAY(memb_race_b_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Black"
If ALL_CLIENTS_ARRAY(memb_race_n_checkbox, 0) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
If ALL_CLIENTS_ARRAY(memb_race_p_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
If ALL_CLIENTS_ARRAY(memb_race_w_checkbox, 0) = checked then race_to_enter = race_to_enter & ", White"
If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

objPers1Table.Cell(16, 1).Range.Text = progs_applying_for
objPers1Table.Cell(16, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_ethnicity, 0)
objPers1Table.Cell(16, 3).Range.Text = race_to_enter


objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.TypeText "NOTES: " & ALL_CLIENTS_ARRAY(memb_notes, 0) & vbCR
objSelection.Font.Bold = TRUE
objSelection.TypeText "AGENCY USE:" & vbCr
objSelection.Font.Bold = FALSE
objSelection.TypeText chr(9) & "Intends to reside in MN? - " & ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, 0) & vbCr
objSelection.TypeText chr(9) & "Has Sponsor? - " & ALL_CLIENTS_ARRAY(clt_sponsor_yn, 0) & vbCr
objSelection.TypeText chr(9) & "Immigration Status: " & ALL_CLIENTS_ARRAY(clt_imig_status, 0) & vbCr
objSelection.TypeText chr(9) & "Verification: " & ALL_CLIENTS_ARRAY(clt_verif_yn, 0) & vbCr
If ALL_CLIENTS_ARRAY(clt_verif_details, 0) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & ALL_CLIENTS_ARRAY(clt_verif_details, 0) & vbCr


objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF 1 - EXPEDITED QUESTIONS"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
set objEXPTable = objDoc.Tables(2)		'Creates the table with the specific index'

objEXPTable.AutoFormat(16)							'This adds the borders to the table and formats it
objEXPTable.Columns(1).Width = 375
objEXPTable.Columns(2).Width = 120
' objEXPTable.Columns(1).Font.Bold = TRUE
' objEXPTable.Columns(2).Font.Bold = TRUE
for col = 1 to 2
	for row = 1 to 8
		objEXPTable.Cell(row, col).Range.Font.Bold = TRUE
	next
next

objEXPTable.Cell(1, 1).Range.Text = "1. How much income (cash or checkes) did or will your household get this month?"
objEXPTable.Cell(1, 2).Range.Text = exp_q_1_income_this_month

objEXPTable.Cell(2, 1).Range.Text = "2. How much does your household (including children) have cash, checking or savings?"
objEXPTable.Cell(2, 2).Range.Text = exp_q_2_assets_this_month

objEXPTable.Cell(3, 1).Range.Text = "3. How much does your household pay for rent/mortgage per month?"
objEXPTable.Cell(3, 2).Range.Text = exp_q_3_rent_this_month

objEXPTable.Cell(4, 1).Range.Text = "   What utilities do you pay?"
If exp_pay_heat_checkbox = checked Then util_pay = util_pay & "Heat, "
If exp_pay_ac_checkbox = checked Then util_pay = util_pay & "Air Conditioning, "
If exp_pay_electricity_checkbox = checked Then util_pay = util_pay & "Electricity, "
If exp_pay_phone_checkbox = checked Then util_pay = util_pay & "Phone, "
If exp_pay_none_checkbox = checked Then util_pay = util_pay & "NONE"
If right(util_pay, 2) = ", " Then util_pay = left(util_pay, len(util_pay) - 2)
objEXPTable.Cell(4, 2).Range.Text = util_pay

objEXPTable.Cell(5, 1).Range.Text = "4. Is anyone in your household a migrant or seasonal farm worker?"
objEXPTable.Cell(5, 2).Range.Text = exp_migrant_seasonal_formworker_yn

objEXPTable.Cell(6, 1).Range.Text = "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
objEXPTable.Cell(6, 2).Range.Text = exp_received_previous_assistance_yn

' objEXPTable.Rows(7).Cells.Merge
objEXPTable.Rows(7).Cells.Split 1, 6, TRUE
' objEXPTable.Cell(7, 1).Range.Split 1, 4, FALSE
' objEXPTable.Cell(7, 5).Range.Split 1, 2, FALSE
objEXPTable.Cell(7, 1).Range.Text = "When?"
objEXPTable.Cell(7, 2).Range.Text = exp_previous_assistance_when
objEXPTable.Cell(7, 3).Range.Text = "Where?"
objEXPTable.Cell(7, 4).Range.Text = exp_previous_assistance_where
objEXPTable.Cell(7, 5).Range.Text = "What?"
objEXPTable.Cell(7, 6).Range.Text = exp_previous_assistance_what

objEXPTable.Cell(8, 1).Range.Text = "6. Is anyone in your household pregnant?"
objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_who


objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information
objSelection.Font.Bold = FALSE

table_count = 3
If UBound(ALL_CLIENTS_ARRAY, 2) <> 0 Then
	ReDim TABLE_ARRAY(UBound(ALL_CLIENTS_ARRAY, 2)-1)
	array_counters = 0

	For each_member = 1 to UBound(ALL_CLIENTS_ARRAY, 2)
		objSelection.TypeText "PERSON " & each_member + 1
		Set objRange = objSelection.Range					'range is needed to create tables
		objDoc.Tables.Add objRange, 10, 1					'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		table_count = table_count + 1

		TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 500

		for row = 1 to 9 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
		Next
		for row = 2 to 10 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
		Next

		For row = 1 to 2
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 140, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 85, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 85, 2
			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
		Next
		For col = 1 to 4
			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
		Next

		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "LEGAL NAME - LAST"
		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "FIRST NAME"
		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "MIDDLE NAME"
		TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "OTHER NAMES"

		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_last_name, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_first_name, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_mid_name, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = ALL_CLIENTS_ARRAY(memb_other_names, each_member)

		' TABLE_ARRAY(array_counters).Cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
		' TABLE_ARRAY(array_counters).Cell(1, 3).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
		' TABLE_ARRAY(array_counters).Cell(1, 4).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
		' TABLE_ARRAY(array_counters).Cell(1, 1).Range.Borders(9).LineStyle = 0
		' TABLE_ARRAY(array_counters).Rows(1).Range.Borders(9).LineStyle = 0
		' TABLE_ARRAY(array_counters).Rows(1).Borders(wdBorderBottom) = wdLineStyleNone

		' TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 5, TRUE
		For row = 3 to 4
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 110, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 85, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 115, 2
			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
		Next
		For col = 1 to 4
			TABLE_ARRAY(array_counters).Cell(3, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(4, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
		TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "DATE OF BIRTH"
		TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "GENDER"
		TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "MARITAL STATUS"

		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_soc_sec_numb, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_dob, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_gender, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = ALL_CLIENTS_ARRAY(memi_marriage_status, each_member)


		' TABLE_ARRAY(array_counters).Rows(7).Cells.Split 1, 3, TRUE
		For row = 5 to 6
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 190, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 190, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(5, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(6, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
		TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "WHAT IS YOU RPREFERRED SPOKEN LANGUAGE?"
		TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_interpreter, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_spoken_language, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_written_language, each_member)


		' TABLE_ARRAY(array_counters).Rows(8).Cells.Split 1, 3, TRUE
		For row = 7 to 8
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 270, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(7, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(8, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
		TABLE_ARRAY(array_counters).Cell(7, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
		TABLE_ARRAY(array_counters).Cell(7, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = ALL_CLIENTS_ARRAY(memi_last_grade, each_member)
		TABLE_ARRAY(array_counters).Cell(8, 2).Range.Text = "Date: " & ALL_CLIENTS_ARRAY(memi_MN_entry_date, each_member) & "   From: " & ALL_CLIENTS_ARRAY(memi_former_state, each_member)
		TABLE_ARRAY(array_counters).Cell(8, 3).Range.Text = ALL_CLIENTS_ARRAY(memi_citizen, each_member)

		For row = 9 to 10
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 275, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 95, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 130, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(9, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(10, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(9, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
		TABLE_ARRAY(array_counters).Cell(9, 2).Range.Text = "ETHNICITY"
		TABLE_ARRAY(array_counters).Cell(9, 3).Range.Text = "RACE"

		If ALL_CLIENTS_ARRAY(clt_none_checkbox, each_member) = checked then progs_applying_for = "NONE"
		If ALL_CLIENTS_ARRAY(clt_snap_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", SNAP"
		If ALL_CLIENTS_ARRAY(clt_cash_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Cash"
		If ALL_CLIENTS_ARRAY(clt_emer_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
		If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

		If ALL_CLIENTS_ARRAY(memb_race_a_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Asian"
		If ALL_CLIENTS_ARRAY(memb_race_b_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Black"
		If ALL_CLIENTS_ARRAY(memb_race_n_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
		If ALL_CLIENTS_ARRAY(memb_race_p_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
		If ALL_CLIENTS_ARRAY(memb_race_w_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", White"
		If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

		TABLE_ARRAY(array_counters).Cell(10, 1).Range.Text = progs_applying_for
		TABLE_ARRAY(array_counters).Cell(10, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_ethnicity, each_member)
		TABLE_ARRAY(array_counters).Cell(10, 3).Range.Text = race_to_enter


		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		objSelection.TypeParagraph()						'adds a line between the table and the next information

		objSelection.TypeText "NOTES: " & ALL_CLIENTS_ARRAY(memb_notes, each_member) & vbCR
		objSelection.Font.Bold = TRUE
		objSelection.TypeText "AGENCY USE:" & vbCr
		objSelection.Font.Bold = FALSE
		objSelection.TypeText chr(9) & "Intends to reside in MN? - " & ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, each_member) & vbCr
		objSelection.TypeText chr(9) & "Has Sponsor? - " & ALL_CLIENTS_ARRAY(clt_sponsor_yn, each_member) & vbCr
		objSelection.TypeText chr(9) & "Immigration Status: " & ALL_CLIENTS_ARRAY(clt_imig_status, each_member) & vbCr
		objSelection.TypeText chr(9) & "Verification: " & ALL_CLIENTS_ARRAY(clt_verif_yn, each_member) & vbCr
		If ALL_CLIENTS_ARRAY(clt_verif_details, each_member) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & ALL_CLIENTS_ARRAY(clt_verif_details, each_member) & vbCr

		array_counters = array_counters + 1
	Next
Else
	objSelection.TypeText "THERE ARE NO OTHER PEOPLE TO BE LISTED ON THIS APPLICATION"
End If
Call script_end_procedure("Done")
