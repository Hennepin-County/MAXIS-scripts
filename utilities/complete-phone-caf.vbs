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

caf_form_date = date & ""
Call find_user_name(worker_name)						'defaulting the name of the suer running the script

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

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
Call restore_your_work(vars_filled)

If vars_filled = FALSE Then
