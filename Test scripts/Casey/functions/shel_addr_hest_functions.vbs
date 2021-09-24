'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - HOUSING DETAIL UPDATE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 125          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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


function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_information, update_attempted)
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
		resi_street_full = trim(resi_line_one & " " & resi_line_two)
        resi_city = replace(city_line, "_", "")
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

		Call get_state_name_from_state_code(state_line, resi_state, TRUE)		'This function makes the state code to be the state name written out - including the code

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
        If verif_line = "NO" Then addr_verif = "NO - No Verif"
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
		mail_street_full = trim(mail_line_one & " " & mail_line_two)
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

		original_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							   addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							   mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&verif_received
    End If

    If access_type = "WRITE" Then
		' verif_received 'add functionality to change how this is updated based on if we have verif or not.

		update_attempted = False
		resi_line_one = trim(resi_line_one)
		resi_line_two = trim(resi_line_two)
		resi_street_full = trim(resi_street_full)
		resi_city = trim(resi_city)
		resi_state = trim(resi_state)
		resi_zip = trim(resi_zip)
		resi_county = trim(resi_county)
		addr_verif = trim(addr_verif)
		addr_homeless = trim(addr_homeless)
		addr_reservation = trim(addr_reservation)
		addr_living_sit = trim(addr_living_sit)
		mail_line_one = trim(mail_line_one)
		mail_line_two = trim(mail_line_two)
		mail_street_full = trim(mail_street_full)
		mail_city = trim(mail_city)
		mail_state = trim(mail_state)
		mail_zip = trim(mail_zip)
		addr_eff_date = trim(addr_eff_date)
		addr_future_date = trim(addr_future_date)
		phone_one = trim(phone_one)
		phone_two = trim(phone_two)
		phone_three = trim(phone_three)
		type_one = trim(type_one)
		type_two = trim(type_two)
		type_three = trim(type_three)
		verif_received = trim(verif_received)

		current_information = resi_line_one&"|"&resi_line_two&"|"&resi_street_full&"|"&resi_city&"|"&resi_state&"|"&resi_zip&"|"&resi_county&"|"&addr_verif&"|"&_
							  addr_homeless&"|"&addr_reservation&"|"&addr_living_sit&"|"&mail_line_one&"|"&mail_line_two&"|"&mail_street_full&"|"&mail_city&"|"&_
							  mail_state&"|"&mail_zip&"|"&addr_eff_date&"|"&addr_future_date&"|"&phone_one&"|"&phone_two&"|"&phone_three&"|"&type_one&"|"&type_two&"|"&type_three&"|"&verif_received


		' MsgBox "THIS" & vbCR & "ORIGINAL" & vbCr & original_information & vbCr & vbCr & "CURRENT" & vbCr & current_information
		If current_information <> original_information Then
			update_attempted = True

	        Call navigate_to_MAXIS_screen("STAT", "ADDR")

	        PF9

	        Call create_mainframe_friendly_date(addr_eff_date, 4, 43, "YY")

			If resi_street_full <> "" AND resi_line_one = "" Then resi_line_one = resi_street_full
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
	        EMWriteScreen left(resi_county, 2), 9, 66
	        EMWriteScreen left(resi_state, 2), 8, 66
	        EMWriteScreen resi_zip, 9, 43

	        EMWriteScreen left(addr_verif, 2), 9, 74

			If mail_street_full <> "" AND mail_line_one = "" Then mail_line_one = mail_street_full
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

			If mail_street_full = "" Then
				EMSetCursor 13, 43
				EMSendKey "<eraseEOF>"

				EMSetCursor 14, 43
				EMSendKey "<eraseEOF>"

				EMSetCursor 16, 43
				EMSendKey "<eraseEOF>"
			End If
			If write_mail_line_one = "" Then
				EMSetCursor 13, 43
				EMSendKey "<eraseEOF>"

				EMSetCursor 16, 43
				EMSendKey "<eraseEOF>"
			End If
			If write_mail_line_two = "" Then
				EMSetCursor 14, 43
				EMSendKey "<eraseEOF>"
			End If
			If mail_city = "" Then
				EMSetCursor 15, 43
				EMSendKey "<eraseEOF>"
			End If
			If mail_zip = "" Then
				EMSetCursor 16, 52
				EMSendKey "<eraseEOF>"
			End If

	        call split_phone_number_into_parts(phone_one, phone_one_left, phone_one_mid, phone_one_right)
	        call split_phone_number_into_parts(phone_two, phone_two_left, phone_two_mid, phone_two_right)
	        call split_phone_number_into_parts(phone_three, phone_three_left, phone_three_mid, phone_three_right)

	        EMWriteScreen phone_one_left, 17, 45
	        EMWriteScreen phone_one_mid, 17, 51
	        EMWriteScreen phone_one_right, 17, 55
	        If type_one <> "Select One..." Then EMWriteScreen type_one, 17, 67

	        EMWriteScreen phone_two_left, 18, 45
	        EMWriteScreen phone_two_mid, 18, 51
	        EMWriteScreen phone_two_right, 18, 55
	        If type_two <> "Select One..." Then EMWriteScreen type_two, 18, 67

	        EMWriteScreen phone_three_left, 19, 45
	        EMWriteScreen phone_three_mid, 19, 51
	        EMWriteScreen phone_three_right, 19, 55
	        If type_three <> "Select One..." Then EMWriteScreen type_three, 19, 67

	        save_attempt = 1
	        Do
	            transmit
				MsgBox "Pause - " & save_attempt
	            EMReadScreen resi_standard_note, 33, 24, 2
	            If resi_standard_note = "RESIDENCE ADDRESS IS STANDARDIZED" Then transmit
	            EMReadScreen mail_standard_note, 31, 24, 2
	            If mail_standard_note = "MAILING ADDRESS IS STANDARDIZED" Then transmit

				EMReadScreen warn_msg, 60, 24, 2
				warn_msg = trim(warn_msg)
				If warn_msg = "ENTER A VALID COMMAND OR PF-KEY" Then Exit Do
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


	        ' Dialog1 = ""
	        ' BeginDialog Dialog1, 0, 0, 356, 160, "ADDR Updated"
	        '   EditBox 60, 120, 290, 15, notes_on_address
	        '   ButtonGroup ButtonPressed
	        '     OkButton 300, 140, 50, 15
	        '   Text 10, 10, 160, 10, "The ADDR panel has been updated successfully. "
	        '   Text 10, 30, 155, 20, "When saving the information to the panel, the following warning message was displayed:"
	        '   Text 30, 55, 310, 55, warning_message
	        '   Text 5, 125, 50, 10, "Address Notes:"
	        ' EndDialog
			'
	        ' Do
	        '     err_msg = ""
	        '     dialog Dialog1
	        '     cancel_confirmation
			'
	        '     EMReadScreen addr_check, 4, 2, 44
	        '     If addr_check = "ADDR" Then
	        '         EMReadScreen info_saved, 7, 24, 2
	        '         If info_saved <> "ENTER A"  Then err_msg = err_msg & vbNewLine & "* Review the ADDR panel and update as needed. It appears the script is unable to complete the update without assistance. In order to prevent all work from being lost, please complete the ADDR update manually and press 'OK' for the script to continue once the address information has been saved."
	        '     End If
			'
	        '     If err_msg <> "" Then MsgBox "The ADDR Update functionality needs assistance" & vbNewLine & err_msg
	        ' Loop until err_msg = ""
		End IF
    End If
end function


	'READ and WRITE SHEL - verif and not - handle for MEMBERS

function access_SHEL_panel(access_type, shel_ref_number, hud_sub_yn, shared_yn, paid_to, rent_retro_amt, rent_retro_verif, rent_prosp_amt, rent_prosp_verif, lot_rent_retro_amt, lot_rent_retro_verif, lot_rent_prosp_amt, lot_rent_prosp_verif, mortgage_retro_amt, mortgage_retro_verif, mortgage_prosp_amt, mortgage_prosp_verif, insurance_retro_amt, insurance_retro_verif, insurance_prosp_amt, insurance_prosp_verif, tax_retro_amt, tax_retro_verif, tax_prosp_amt, tax_prosp_verif, room_retro_amt, room_retro_verif, room_prosp_amt, room_prosp_verif, garage_retro_amt, garage_retro_verif, garage_prosp_amt, garage_prosp_verif, subsidy_retro_amt, subsidy_retro_verif, subsidy_prosp_amt, subsidy_prosp_verif, original_information)
	Call navigate_to_MAXIS_screen("STAT", "SHEL")
	EMWriteScreen shel_ref_number, 20, 76
	transmit

	access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen hud_sub_yn,            1, 6, 46
        EMReadScreen shared_yn,             1, 6, 64
        EMReadScreen paid_to,               25, 7, 50

        paid_to = replace(paid_to, "_", "")

        EMReadScreen rent_retro_amt,        8, 11, 37
        EMReadScreen rent_retro_verif,      2, 11, 48
        EMReadScreen rent_prosp_amt,        8, 11, 56
        EMReadScreen rent_prosp_verif,      2, 11, 67

        rent_retro_amt = replace(rent_retro_amt, "_", "")
        rent_retro_amt = trim(rent_retro_amt)
        If rent_retro_verif = "SF" Then rent_retro_verif = "SF - Shelter Form"
        If rent_retro_verif = "LE" Then rent_retro_verif = "LE - Lease"
        If rent_retro_verif = "RE" Then rent_retro_verif = "RE - Rent Receipt"
        If rent_retro_verif = "OT" Then rent_retro_verif = "OT - Other Doc"
        If rent_retro_verif = "NC" Then rent_retro_verif = "NC - Chg, Neg Impact"
        If rent_retro_verif = "PC" Then rent_retro_verif = "PC - Chg, Pos Impact"
        If rent_retro_verif = "NO" Then rent_retro_verif = "NO - No Verif"
		If rent_retro_verif = "?_" Then rent_retro_verif = "? - Delayed Verif"
        If rent_retro_verif = "__" Then rent_retro_verif = ""
        rent_prosp_amt = replace(rent_prosp_amt, "_", "")
        rent_prosp_amt = trim(rent_prosp_amt)
		If rent_prosp_amt = "" Then rent_prosp_amt = 0
		rent_prosp_amt = rent_prosp_amt * 1
        If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
        If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
        If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
        If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Doc"
        If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg, Neg Impact"
        If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg, Pos Impact"
        If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Verif"
		If rent_prosp_verif = "?_" Then rent_prosp_verif = "? - Delayed Verif"
        If rent_prosp_verif = "__" Then rent_prosp_verif = ""

        EMReadScreen lot_rent_retro_amt,    8, 12, 37
        EMReadScreen lot_rent_retro_verif,  2, 12, 48
        EMReadScreen lot_rent_prosp_amt,    8, 12, 56
        EMReadScreen lot_rent_prosp_verif,  2, 12, 67

        lot_rent_retro_amt = replace(lot_rent_retro_amt, "_", "")
        lot_rent_retro_amt = trim(lot_rent_retro_amt)
        If lot_rent_retro_verif = "LE" Then lot_rent_retro_verif = "LE - Lease"
        If lot_rent_retro_verif = "RE" Then lot_rent_retro_verif = "RE - Rent Receipt"
        If lot_rent_retro_verif = "BI" Then lot_rent_retro_verif = "BI - Billing Stmt"
        If lot_rent_retro_verif = "OT" Then lot_rent_retro_verif = "OT - Other Doc"
        If lot_rent_retro_verif = "NC" Then lot_rent_retro_verif = "NC - Chg, Neg Impact"
        If lot_rent_retro_verif = "PC" Then lot_rent_retro_verif = "PC - Chg, Pos Impact"
        If lot_rent_retro_verif = "NO" Then lot_rent_retro_verif = "NO - No Verif"
		If lot_rent_retro_verif = "?_" Then lot_rent_retro_verif = "? - Delayed Verif"
        If lot_rent_retro_verif = "__" Then lot_rent_retro_verif = ""
        lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
        lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
		If lot_rent_prosp_amt = "" Then lot_rent_prosp_amt = 0
		lot_rent_prosp_amt = lot_rent_prosp_amt * 1
        If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
        If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
        If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
        If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Doc"
        If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg, Neg Impact"
        If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg, Pos Impact"
        If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Verif"
		If lot_rent_prosp_verif = "?_" Then lot_rent_prosp_verif = "? - Delayed Verif"
        If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""

        EMReadScreen mortgage_retro_amt,    8, 13, 37
        EMReadScreen mortgage_retro_verif,  2, 13, 48
        EMReadScreen mortgage_prosp_amt,    8, 13, 56
        EMReadScreen mortgage_prosp_verif,  2, 13, 67

        mortgage_retro_amt = replace(mortgage_retro_amt, "_", "")
        mortgage_retro_amt = trim(mortgage_retro_amt)
        If mortgage_retro_verif = "MO" Then mortgage_retro_verif = "MO - Mortgage Pmt Book"
        If mortgage_retro_verif = "CD" Then mortgage_retro_verif = "CD - Ctrct fro Deed"
        If mortgage_retro_verif = "OT" Then mortgage_retro_verif = "OT - Other Doc"
        If mortgage_retro_verif = "NC" Then mortgage_retro_verif = "NC - Chg, Neg Impact"
        If mortgage_retro_verif = "PC" Then mortgage_retro_verif = "PC - Chg, Pos Impact"
        If mortgage_retro_verif = "NO" Then mortgage_retro_verif = "NO - No Verif"
		If mortgage_retro_verif = "?_" Then mortgage_retro_verif = "? - Delayed Verif"
        If mortgage_retro_verif = "__" Then mortgage_retro_verif = ""
        mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
        mortgage_prosp_amt = trim(mortgage_prosp_amt)
		If mortgage_prosp_amt = "" Then mortgage_prosp_amt = 0
		mortgage_prosp_amt = mortgage_prosp_amt * 1
        If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
        If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
        If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Doc"
        If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg, Neg Impact"
        If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg, Pos Impact"
        If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Verif"
		If mortgage_prosp_verif = "?_" Then mortgage_prosp_verif = "? - Delayed Verif"
        If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""

        EMReadScreen insurance_retro_amt,   8, 14, 37
        EMReadScreen insurance_retro_verif, 2, 14, 48
        EMReadScreen insurance_prosp_amt,   8, 14, 56
        EMReadScreen insurance_prosp_verif, 2, 14, 67

        insurance_retro_amt = replace(insurance_retro_amt, "_", "")
        insurance_retro_amt = trim(insurance_retro_amt)
        If insurance_retro_verif = "BI" Then insurance_retro_verif = "BI - Billing Stmt"
        If insurance_retro_verif = "OT" Then insurance_retro_verif = "OT - Other Doc"
        If insurance_retro_verif = "NC" Then insurance_retro_verif = "NC - Chg, Neg Impact"
        If insurance_retro_verif = "PC" Then insurance_retro_verif = "PC - Chg, Pos Impact"
        If insurance_retro_verif = "NO" Then insurance_retro_verif = "NO - No Verif"
		If insurance_retro_verif = "?_" Then insurance_retro_verif = "? - Delayed Verif"
        If insurance_retro_verif = "__" Then insurance_retro_verif = ""
        insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
        insurance_prosp_amt = trim(insurance_prosp_amt)
		If insurance_prosp_amt = "" Then insurance_prosp_amt = 0
		insurance_prosp_amt = insurance_prosp_amt * 1
        If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
        If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Doc"
        If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg, Neg Impact"
        If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg, Pos Impact"
        If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Verif"
		If insurance_prosp_verif = "?_" Then insurance_prosp_verif = "? - Delayed Verif"
        If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""

        EMReadScreen tax_retro_amt,         8, 15, 37
        EMReadScreen tax_retro_verif,       2, 15, 48
        EMReadScreen tax_prosp_amt,         8, 15, 56
        EMReadScreen tax_prosp_verif,       2, 15, 67

        tax_retro_amt = replace(tax_retro_amt, "_", "")
        tax_retro_amt = trim(tax_retro_amt)
        If tax_retro_verif = "TX" Then tax_retro_verif = "TX - Prop Tax Stmt"
        If tax_retro_verif = "OT" Then tax_retro_verif = "OT - Other Doc"
        If tax_retro_verif = "NC" Then tax_retro_verif = "NC - Chg, Neg Impact"
        If tax_retro_verif = "PC" Then tax_retro_verif = "PC - Chg, Pos Impact"
        If tax_retro_verif = "NO" Then tax_retro_verif = "NO - No Verif"
		If tax_retro_verif = "?_" Then tax_retro_verif = "? - Delayed Verif"
        If tax_retro_verif = "__" Then tax_retro_verif = ""
        tax_prosp_amt = replace(tax_prosp_amt, "_", "")
        tax_prosp_amt = trim(tax_prosp_amt)
		If tax_prosp_amt = "" Then tax_prosp_amt = 0
		tax_prosp_amt = tax_prosp_amt * 1
        If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
        If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Doc"
        If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg, Neg Impact"
        If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg, Pos Impact"
        If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Verif"
		If tax_prosp_verif = "?_" Then tax_prosp_verif = "? - Delayed Verif"
        If tax_prosp_verif = "__" Then tax_prosp_verif = ""

        EMReadScreen room_retro_amt,        8, 16, 37
        EMReadScreen room_retro_verif,      2, 16, 48
        EMReadScreen room_prosp_amt,        8, 16, 56
        EMReadScreen room_prosp_verif,      2, 16, 67

        room_retro_amt = replace(room_retro_amt, "_", "")
        room_retro_amt = trim(room_retro_amt)
        If room_retro_verif = "SF" Then room_retro_verif = "SF - Shelter Form"
        If room_retro_verif = "LE" Then room_retro_verif = "LE - Lease"
        If room_retro_verif = "RE" Then room_retro_verif = "RE - Rent Receipt"
        If room_retro_verif = "OT" Then room_retro_verif = "OT - Other Doc"
        If room_retro_verif = "NC" Then room_retro_verif = "NC - Chg, Neg Impact"
        If room_retro_verif = "PC" Then room_retro_verif = "PC - Chg, Pos Impact"
        If room_retro_verif = "NO" Then room_retro_verif = "NO - No Verif"
		If room_retro_verif = "?_" Then room_retro_verif = "? - Delayed Verif"
        If room_retro_verif = "__" Then room_retro_verif = ""
        room_prosp_amt = replace(room_prosp_amt, "_", "")
        room_prosp_amt = trim(room_prosp_amt)
		If room_prosp_amt = "" Then room_prosp_amt = 0
		room_prosp_amt = room_prosp_amt * 1
        If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
        If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
        If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
        If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Doc"
        If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg, Neg Impact"
        If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg, Pos Impact"
        If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Verif"
		If room_prosp_verif = "?_" Then room_prosp_verif = "? - Delayed Verif"
        If room_prosp_verif = "__" Then room_prosp_verif = ""

        EMReadScreen garage_retro_amt,      8, 17, 37
        EMReadScreen garage_retro_verif,    2, 17, 48
        EMReadScreen garage_prosp_amt,      8, 17, 56
        EMReadScreen garage_prosp_verif,    2, 17, 67

        garage_retro_amt = replace(garage_retro_amt, "_", "")
        garage_retro_amt = trim(garage_retro_amt)
        If garage_retro_verif = "SF" Then garage_retro_verif = "SF - Shelter Form"
        If garage_retro_verif = "LE" Then garage_retro_verif = "LE - Lease"
        If garage_retro_verif = "RE" Then garage_retro_verif = "RE - Rent Receipt"
        If garage_retro_verif = "OT" Then garage_retro_verif = "OT - Other Doc"
        If garage_retro_verif = "NC" Then garage_retro_verif = "NC - Chg, Neg Impact"
        If garage_retro_verif = "PC" Then garage_retro_verif = "PC - Chg, Pos Impact"
        If garage_retro_verif = "NO" Then garage_retro_verif = "NO - No Verif"
		If garage_retro_verif = "?_" Then garage_retro_verif = "? - Delayed Verif"
        If garage_retro_verif = "__" Then garage_retro_verif = ""
        garage_prosp_amt = replace(garage_prosp_amt, "_", "")
        garage_prosp_amt = trim(garage_prosp_amt)
		If garage_prosp_amt = "" Then garage_prosp_amt = 0
		garage_prosp_amt = garage_prosp_amt * 1
        If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
        If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
        If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
        If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Doc"
        If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg, Neg Impact"
        If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg, Pos Impact"
        If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Verif"
		If garage_prosp_verif = "?_" Then garage_prosp_verif = "? - Delayed Verif"
        If garage_prosp_verif = "__" Then garage_prosp_verif = ""

        EMReadScreen subsidy_retro_amt,     8, 18, 37
        EMReadScreen subsidy_retro_verif,   2, 18, 48
        EMReadScreen subsidy_prosp_amt,     8, 18, 56
        EMReadScreen subsidy_prosp_verif,   2, 18, 67

        subsidy_retro_amt = replace(subsidy_retro_amt, "_", "")
        subsidy_retro_amt = trim(subsidy_retro_amt)
        If subsidy_retro_verif = "SF" Then subsidy_retro_verif = "SF - Shelter Form"
        If subsidy_retro_verif = "LE" Then subsidy_retro_verif = "LE - Lease"
        If subsidy_retro_verif = "OT" Then subsidy_retro_verif = "OT - Other Doc"
        If subsidy_retro_verif = "NO" Then subsidy_retro_verif = "NO - No Verif"
		If subsidy_retro_verif = "?_" Then subsidy_retro_verif = "? - Delayed Verif"
        If subsidy_retro_verif = "__" Then subsidy_retro_verif = ""
        subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
        subsidy_prosp_amt = trim(subsidy_prosp_amt)
		If subsidy_prosp_amt = "" Then subsidy_prosp_amt = 0
		subsidy_prosp_amt = subsidy_prosp_amt * 1
        If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
        If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
        If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Doc"
        If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Verif"
		If subsidy_prosp_verif = "?_" Then subsidy_prosp_verif = "? - Delayed Verif"
        If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""

		original_information = hud_sub_yn&"|"&shared_yn&"|"&paid_to&"|"&rent_retro_amt&"|"&rent_retro_verif&"|"&rent_prosp_amt&"|"&rent_prosp_verif&"|"&lot_rent_retro_amt&"|"&lot_rent_retro_verif&"|"&lot_rent_prosp_amt&"|"&_
							   lot_rent_prosp_verif&"|"&mortgage_retro_amt&"|"&mortgage_retro_verif&"|"&mortgage_prosp_amt&"|"&mortgage_prosp_verif&"|"&insurance_retro_amt&"|"&insurance_retro_verif&"|"&insurance_prosp_amt&"|"&_
							   insurance_prosp_verif&"|"&tax_retro_amt&"|"&tax_retro_verif&"|"&tax_prosp_amt&"|"&tax_prosp_verif&"|"&room_retro_amt&"|"&room_retro_verif&"|"&room_prosp_amt&"|"&room_prosp_verif&"|"&garage_retro_amt&"|"&_
							   garage_retro_verif&"|"&garage_prosp_amt&"|"&garage_prosp_verif&"|"&subsidy_retro_amt&"|"&subsidy_retro_verif&"|"&subsidy_prosp_amt&"|"&subsidy_prosp_verif
    End If

	If access_type = "WRITE" Then
		hud_sub_yn = trim(hud_sub_yn)
		shared_yn = trim(shared_yn)
		paid_to = trim(paid_to)
		rent_retro_amt = trim(rent_retro_amt)
		rent_retro_verif = trim(rent_retro_verif)
		rent_prosp_amt = trim(rent_prosp_amt)
		rent_prosp_verif = trim(rent_prosp_verif)
		lot_rent_retro_amt = trim(lot_rent_retro_amt)
		lot_rent_retro_verif = trim(lot_rent_retro_verif)
		lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
		lot_rent_prosp_verif = trim(lot_rent_prosp_verif)
		mortgage_retro_amt = trim(mortgage_retro_amt)
		mortgage_retro_verif = trim(mortgage_retro_verif)
		mortgage_prosp_amt = trim(mortgage_prosp_amt)
		mortgage_prosp_verif = trim(mortgage_prosp_verif)
		insurance_retro_amt = trim(insurance_retro_amt)
		insurance_retro_verif = trim(insurance_retro_verif)
		insurance_prosp_amt = trim(insurance_prosp_amt)
		insurance_prosp_verif = trim(insurance_prosp_verif)
		tax_retro_amt = trim(tax_retro_amt)
		tax_retro_verif = trim(tax_retro_verif)
		tax_prosp_amt = trim(tax_prosp_amt)
		tax_prosp_verif = trim(tax_prosp_verif)
		room_retro_amt = trim(room_retro_amt)
		room_retro_verif = trim(room_retro_verif)
		room_prosp_amt = trim(room_prosp_amt)
		room_prosp_verif = trim(room_prosp_verif)
		garage_retro_amt = trim(garage_retro_amt)
		garage_retro_verif = trim(garage_retro_verif)
		garage_prosp_amt = trim(garage_prosp_amt)
		garage_prosp_verif = trim(garage_prosp_verif)
		subsidy_retro_amt = trim(subsidy_retro_amt)
		subsidy_retro_verif = trim(subsidy_retro_verif)
		subsidy_prosp_amt = trim(subsidy_prosp_amt)
		subsidy_prosp_verif = trim(subsidy_prosp_verif)

		current_shel_details = hud_sub_yn&"|"&shared_yn&"|"&paid_to&"|"&rent_retro_amt&"|"&rent_retro_verif&"|"&rent_prosp_amt&"|"&rent_prosp_verif&"|"&lot_rent_retro_amt&"|"&lot_rent_retro_verif&"|"&lot_rent_prosp_amt&"|"&_
							   lot_rent_prosp_verif&"|"&mortgage_retro_amt&"|"&mortgage_retro_verif&"|"&mortgage_prosp_amt&"|"&mortgage_prosp_verif&"|"&insurance_retro_amt&"|"&insurance_retro_verif&"|"&insurance_prosp_amt&"|"&_
							   insurance_prosp_verif&"|"&tax_retro_amt&"|"&tax_retro_verif&"|"&tax_prosp_amt&"|"&tax_prosp_verif&"|"&room_retro_amt&"|"&room_retro_verif&"|"&room_prosp_amt&"|"&room_prosp_verif&"|"&garage_retro_amt&"|"&_
							   garage_retro_verif&"|"&garage_prosp_amt&"|"&garage_prosp_verif&"|"&subsidy_retro_amt&"|"&subsidy_retro_verif&"|"&subsidy_prosp_amt&"|"&subsidy_prosp_verif



		If current_shel_details <> original_information Then
			EMReadScreen shel_version, 1, 2, 73
			If shel_version = "1" Then PF9
			If shel_version = "0" Then
				EMWriteScreen "nn", 20, 79
				transmit
			End If

			rent_retro_verif = 		rent_retro_verif & "  "
			rent_prosp_verif = 		rent_prosp_verif & "  "
			lot_rent_retro_verif = 	lot_rent_retro_verif & "  "
			lot_rent_prosp_verif = 	lot_rent_prosp_verif & "  "
			mortgage_retro_verif = 	mortgage_retro_verif & "  "
			mortgage_prosp_verif = 	mortgage_prosp_verif & "  "
			insurance_retro_verif = insurance_retro_verif & "  "
			insurance_prosp_verif = insurance_prosp_verif & "  "
			tax_retro_verif = 		tax_retro_verif & "  "
			tax_prosp_verif = 		tax_prosp_verif & "  "
			room_retro_verif = 		room_retro_verif & "  "
			room_prosp_verif = 		room_prosp_verif & "  "
			garage_retro_verif = 	garage_retro_verif & "  "
			garage_prosp_verif = 	garage_prosp_verif & "  "
			subsidy_retro_verif = 	subsidy_retro_verif & "  "
			subsidy_prosp_verif = 	subsidy_prosp_verif & "  "

			If hud_sub_yn = "" Then
				EMSetCursor 6, 46
				EMSendKey "<eraseEOF>"
			Else
				EMWriteScreen hud_sub_yn, 6, 46
			End If
			If shared_yn = "" Then
				EMSetCursor 6, 64
				EMSendKey "<eraseEOF>"
			Else
	        	EMWriteScreen shared_yn, 6, 64
			End If
			If paid_to = "" Then
				EMSetCursor 7, 50
				EMSendKey "<eraseEOF>"
			Else
	        	EMWriteScreen paid_to, 7, 50
			End If

			EMWriteScreen right("        " & rent_retro_amt, 8), 		11, 37
	        EMWriteScreen left(rent_retro_verif, 2),      				11, 48
	        EMWriteScreen right("        " & rent_prosp_amt, 8),    	11, 56
	        EMWriteScreen left(rent_prosp_verif, 2),      				11, 67

			EMWriteScreen right("        " & lot_rent_retro_amt, 8),    12, 37
	        EMWriteScreen left(lot_rent_retro_verif, 2),  				12, 48
	        EMWriteScreen right("        " & lot_rent_prosp_amt, 8),    12, 56
	        EMWriteScreen left(lot_rent_prosp_verif, 2),  				12, 67

			EMWriteScreen right("        " & mortgage_retro_amt, 8),    13, 37
	        EMWriteScreen left(mortgage_retro_verif, 2),  				13, 48
	        EMWriteScreen right("        " & mortgage_prosp_amt, 8),    13, 56
	        EMWriteScreen left(mortgage_prosp_verif, 2),  				13, 67

			EMWriteScreen right("        " & insurance_retro_amt, 8),   14, 37
	        EMWriteScreen left(insurance_retro_verif, 2), 				14, 48
	        EMWriteScreen right("        " & insurance_prosp_amt, 8),   14, 56
	        EMWriteScreen left(insurance_prosp_verif, 2), 				14, 67

			EMWriteScreen right("        " & tax_retro_amt, 8),         15, 37
	        EMWriteScreen left(tax_retro_verif, 2),       				15, 48
	        EMWriteScreen right("        " & tax_prosp_amt, 8),         15, 56
	        EMWriteScreen left(tax_prosp_verif, 2),       				15, 67

			EMWriteScreen right("        " & room_retro_amt, 8),        16, 37
	        EMWriteScreen left(room_retro_verif, 2),      				16, 48
	        EMWriteScreen right("        " & room_prosp_amt, 8),        16, 56
	        EMWriteScreen left(room_prosp_verif, 2),      				16, 67

			EMWriteScreen right("        " & garage_retro_amt, 8),      17, 37
			EMWriteScreen left(garage_retro_verif, 2),    				17, 48
			EMWriteScreen right("        " & garage_prosp_amt, 8),      17, 56
			EMWriteScreen left(garage_prosp_verif, 2),    				17, 67

			EMWriteScreen right("        " & subsidy_retro_amt, 8),     18, 37
			EMWriteScreen left(subsidy_retro_verif, 2),   				18, 48
			EMWriteScreen right("        " & subsidy_prosp_amt, 8),     18, 56
			EMWriteScreen left(subsidy_prosp_verif, 2),   				18, 67

		End If
	End If
end function

'READ and WRITE HEST
function access_HEST_panel(access_type, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
    access_type = UCase(access_type)
	Call navigate_to_MAXIS_screen("STAT", "HEST")
    If access_type = "READ" Then
        hest_col = 40
        Do
            EMReadScreen pers_paying, 2, 6, hest_col
            If pers_paying <> "__" Then
                all_persons_paying = all_persons_paying & ", " & pers_paying
            Else
                exit do
            End If
            hest_col = hest_col + 3
        Loop until hest_col = 70
        If left(all_persons_paying, 1) = "," Then all_persons_paying = right(all_persons_paying, len(all_persons_paying) - 2)

        EMReadScreen choice_date, 8, 7, 40
        EMReadScreen actual_initial_exp, 8, 8, 61

        EMReadScreen retro_heat_ac_yn, 1, 13, 34
        EMReadScreen retro_heat_ac_units, 2, 13, 42
        EMReadScreen retro_heat_ac_amt, 6, 13, 49
        EMReadScreen retro_electric_yn, 1, 14, 34
        EMReadScreen retro_electric_units, 2, 14, 42
        EMReadScreen retro_electric_amt, 6, 14, 49
        EMReadScreen retro_phone_yn, 1, 15, 34
        EMReadScreen retro_phone_units, 2, 15, 42
        EMReadScreen retro_phone_amt, 6, 15, 49

        EMReadScreen prosp_heat_ac_yn, 1, 13, 60
        EMReadScreen prosp_heat_ac_units, 2, 13, 68
        EMReadScreen prosp_heat_ac_amt, 6, 13, 75
        EMReadScreen prosp_electric_yn, 1, 14, 60
        EMReadScreen prosp_electric_units, 2, 14, 68
        EMReadScreen prosp_electric_amt, 6, 14, 75
        EMReadScreen prosp_phone_yn, 1, 15, 60
        EMReadScreen prosp_phone_units, 2, 15, 68
        EMReadScreen prosp_phone_amt, 6, 15, 75

        choice_date = replace(choice_date, " ", "/")
        If choice_date = "__/__/__" Then choice_date = ""
        actual_initial_exp = trim(actual_initial_exp)
        actual_initial_exp = replace(actual_initial_exp, "_", "")

        retro_heat_ac_yn = replace(retro_heat_ac_yn, "_", "")
        retro_heat_ac_units = replace(retro_heat_ac_units, "_", "")
        retro_heat_ac_amt = trim(retro_heat_ac_amt)
		If retro_heat_ac_amt = "" Then retro_heat_ac_amt = 0
		retro_heat_ac_amt = retro_heat_ac_amt * 1
        retro_electric_yn = replace(retro_electric_yn, "_", "")
        retro_electric_units = replace(retro_electric_units, "_", "")
        retro_electric_amt = trim(retro_electric_amt)
		If retro_electric_amt = "" Then retro_electric_amt = 0
		retro_electric_amt = retro_electric_amt * 1
        retro_phone_yn = replace(retro_phone_yn, "_", "")
        retro_phone_units = replace(retro_phone_units, "_", "")
        retro_phone_amt = trim(retro_phone_amt)
		If retro_phone_amt = "" Then retro_phone_amt = 0
		retro_phone_amt = retro_phone_amt * 1

        prosp_heat_ac_yn = replace(prosp_heat_ac_yn, "_", "")
        prosp_heat_ac_units = replace(prosp_heat_ac_units, "_", "")
        prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
        If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
		prosp_heat_ac_amt = prosp_heat_ac_amt * 1
        prosp_electric_yn = replace(prosp_electric_yn, "_", "")
        prosp_electric_units = replace(prosp_electric_units, "_", "")
        prosp_electric_amt = trim(prosp_electric_amt)
        If prosp_electric_amt = "" Then prosp_electric_amt = 0
		prosp_electric_amt = prosp_electric_amt * 1
        prosp_phone_yn = replace(prosp_phone_yn, "_", "")
        prosp_phone_units = replace(prosp_phone_units, "_", "")
        prosp_phone_amt = trim(prosp_phone_amt)
        If prosp_phone_amt = "" Then prosp_phone_amt = 0
		prosp_phone_amt = prosp_phone_amt * 1

        total_utility_expense = 0
        If prosp_heat_ac_yn = "Y" Then
            total_utility_expense =  prosp_heat_ac_amt
        ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt + prosp_phone_amt
        ElseIf prosp_electric_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt
        Elseif prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_phone_amt
        End If

    End If

	If access_type = "WRITE" Then
		EMReadScreen hest_version, 1, 2, 73
		If hest_version = "1" Then PF9
		If hest_version = "0" Then
			EMWriteScreen "nn", 20, 79
			transmit
		End If

		all_persons_paying = trim(all_persons_paying)
		If all_persons_paying <> "" Then
			If InStr(all_persons_paying, ",") = 0 Then
				persons_array = array(all_persons_paying)
			Else
				persons_array = split(all_persons_paying, ",")
			End If

			hest_col = 40
			for each pers_paying in persons_array
				EMWriteScreen pers_paying, 6, hest_col
				hest_col = hest_col + 3
			Next

			If IsDate(choice_date) = True Then Call create_mainframe_friendly_date(choice_date, 7, 40, "YY")
	        EMWriteScreen actual_initial_exp, 8, 61

			EMWriteScreen retro_heat_ac_yn, 13, 34
	        EMWriteScreen retro_heat_ac_units, 13, 42
	        EMWriteScreen retro_electric_yn, 14, 34
	        EMWriteScreen retro_electric_units, 14, 42
	        EMWriteScreen retro_phone_yn, 15, 34
	        EMWriteScreen retro_phone_units, 15, 42

	        EMWriteScreen prosp_heat_ac_yn, 13, 60
	        EMWriteScreen prosp_heat_ac_units, 13, 68
	        EMWriteScreen prosp_electric_yn, 14, 60
	        EMWriteScreen prosp_electric_units, 14, 68
	        EMWriteScreen prosp_phone_yn, 15, 60
	        EMWriteScreen prosp_phone_units, 15, 68


			transmit

			EMReadScreen retro_heat_ac_amt, 6, 13, 49
			EMReadScreen retro_electric_amt, 6, 14, 49
			EMReadScreen retro_phone_amt, 6, 15, 49

			EMReadScreen prosp_heat_ac_amt, 6, 13, 75
			EMReadScreen prosp_electric_amt, 6, 14, 75
			EMReadScreen prosp_phone_amt, 6, 15, 75

			retro_heat_ac_amt = trim(retro_heat_ac_amt)
			If retro_heat_ac_amt = "" Then retro_heat_ac_amt = 0
			retro_heat_ac_amt = retro_heat_ac_amt * 1
			retro_electric_amt = trim(retro_electric_amt)
			If retro_electric_amt = "" Then retro_electric_amt = 0
			retro_electric_amt = retro_electric_amt * 1
			retro_phone_amt = trim(retro_phone_amt)
			If retro_phone_amt = "" Then retro_phone_amt = 0
			retro_phone_amt = retro_phone_amt * 1

			prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
			If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
			prosp_heat_ac_amt = prosp_heat_ac_amt * 1
			prosp_electric_amt = trim(prosp_electric_amt)
			If prosp_electric_amt = "" Then prosp_electric_amt = 0
			prosp_electric_amt = prosp_electric_amt * 1
			prosp_phone_amt = trim(prosp_phone_amt)
			If prosp_phone_amt = "" Then prosp_phone_amt = 0
			prosp_phone_amt = prosp_phone_amt * 1
		End If
	End If
end function

function display_HOUSING_CHANGE_information(housing_questions_step, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn)

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

		Text 15, 25, 450, 10, "STEP 2 - ADDR UPDATES"

		Call display_ADDR_information(True, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)

		' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	End If
	If housing_questions_step = 3 Then
		Text shel_up_x_pos + 5, 10, 60, 10, "SHEL UPDATE"

		Text 15, 25, 450, 10, "STEP 3 - SHEL UPDATES"

		' Text 20, 145, 160, 10, "Have we received verification of this expense?"
		' DropListBox 180, 140, 45, 45, yes_no_list, shel_verif_received_yn

		Text 40, 45, 120, 10, "Is the new expense amount shared?"
		DropListBox 165, 40, 45, 45, yes_no_list, shel_shared_yn
		Text 240, 45, 135, 10, "Is the new expense amount subsidized?"
		DropListBox 380, 40, 45, 45, yes_no_list, shel_subsidized_yn

		EditBox 105, 95, 45, 15, total_current_rent
		DropListBox 155, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_rent_verif
		EditBox 105, 115, 45, 15, total_current_lot_rent
		DropListBox 155, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_lot_rent_verif
		EditBox 105, 135, 45, 15, total_current_mortgage
		DropListBox 155, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_mortgage_verif
		EditBox 105, 155, 45, 15, total_current_insurance
		DropListBox 155, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_insurance_verif
		EditBox 105, 175, 45, 15, total_current_taxes
		DropListBox 155, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_taxes_verif
		EditBox 105, 195, 45, 15, total_current_room
		DropListBox 155, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_room_verif
		EditBox 105, 215, 45, 15, total_current_garage
		DropListBox 155, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Multiple", all_garage_verif
		EditBox 105, 235, 45, 15, total_current_subsidy
		DropListBox 155, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", all_subsidy_verif

		Text 105, 75, 100, 10, "TOTAL Expense Amounts"
	    Text 105, 85, 30, 10, "Amount"
	    Text 160, 85, 20, 10, "Verif"
		Text 80, 100, 20, 10, "Rent:"
	    Text 70, 120, 30, 10, "Lot Rent:"
	    Text 65, 140, 35, 10, "Mortgage:"
	    Text 65, 160, 40, 10, "Insurance:"
	    Text 75, 180, 25, 10, "Taxes:"
	    Text 75, 200, 25, 10, "Room:"
	    Text 75, 220, 30, 10, "Garage:"
	    Text 70, 240, 30, 10, "Subsidy:"


		' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
		' PushButton x_pos + 60, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	End If
	If housing_questions_step = 4 Then
		Text shel_det_x_pos + 5, 10, 60, 10, "SHEL DETAILS"

		Text 15, 25, 450, 10, "STEP 4 - SHEL Details"

		' PushButton x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
		' PushButton x_pos + 60, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
		' PushButton x_pos + 120, 8, 60, 13, "SHEL UPDATE", housing_change_shel_update_btn
	End If
	If housing_questions_step = 5 Then
		PushButton 420, 10, 60, 10, "REVIEW"

		Text 15, 25, 450, 10, "STEP 5 - REVIEW AND CONFIRM"

	End If

	Text 20, 350, 55, 10, "Additional Notes:"
	EditBox 80, 345, 385, 15, addr_or_shel_change_notes

	If housing_questions_step <> 1 Then PushButton overview_x_pos, 8, 60, 13, "OVERVIEW", housing_change_overview_btn
	If view_addr_update_dlg = True AND housing_questions_step <> 2 Then PushButton addr_up_x_pos, 8, 60, 13, "ADDR UPDATE", housing_change_addr_update_btn
	If view_shel_update_dlg = True AND housing_questions_step <> 3 Then PushButton shel_up_x_pos, 8, 60, 13, "SHEL UPDATE", housing_change_shel_update_btn
	If view_shel_details_dlg = True AND housing_questions_step <> 4 Then PushButton shel_det_x_pos, 8, 60, 13, "SHEL DETAILS", housing_change_shel_details_btn
	If err_msg = "" AND housing_questions_step <> 5 Then PushButton 405, 8, 60, 13, "REVIEW", housing_change_review_btn

	If housing_questions_step <> 5 Then PushButton 390, 325, 70, 10, "CONTINUE", housing_change_continue_btn

end function

function display_ADDR_information(update_addr, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)

	If update_addr = False Then
		Text 70, 55, 305, 15, resi_street_full
		Text 70, 75, 105, 15, resi_city
		Text 205, 75, 110, 45, resi_state
		Text 340, 75, 35, 15, resi_zip
		Text 125, 95, 45, 45, addr_reservation
		Text 245, 85, 130, 15, reservation_name
		Text 125, 115, 45, 45, addr_homeless
		If addr_living_sit = "10 - Unknown" OR addr_living_sit = "Blank" Then
			DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
		Else
			Text 245, 115, 130, 45, addr_living_sit
		End If
		Text 70, 165, 305, 15, mail_street_full
		Text 70, 185, 105, 15, mail_city
		Text 205, 185, 110, 45, mail_state
		Text 340, 185, 35, 15, mail_zip
		Text 20, 240, 90, 15, phone_one
		Text 125, 240, 65, 45, type_one
		Text 20, 260, 90, 15, phone_two
		Text 125, 260, 65, 45, type_two
		Text 20, 280, 90, 15, phone_three
		Text 125, 280, 65, 45, type_three
		Text 325, 220, 50, 15, address_change_date
		Text 255, 255, 120, 45, resi_county
		PushButton 290, 300, 95, 15, "Update Information", update_information_btn
	End If
	If update_addr = True Then
		EditBox 70, 50, 305, 15, resi_street_full
		EditBox 70, 70, 105, 15, resi_city
		DropListBox 205, 70, 110, 45, ""+chr(9)+state_list, resi_state
		EditBox 340, 70, 35, 15, resi_zip
		DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", addr_reservation
		EditBox 245, 90, 130, 15, reservation_name
		DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", addr_homeless
		DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", addr_living_sit
		EditBox 70, 160, 305, 15, mail_street_full
		EditBox 70, 180, 105, 15, mail_city
		DropListBox 205, 180, 110, 45, ""+chr(9)+state_list, mail_state
		EditBox 340, 180, 35, 15, mail_zip
		EditBox 20, 240, 90, 15, phone_one
		DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_one
		EditBox 20, 260, 90, 15, phone_two
		DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_two
		EditBox 20, 280, 90, 15, phone_three
		DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", type_three
		EditBox 325, 220, 50, 15, address_change_date
		ComboBox 255, 255, 120, 45, county_list_smalll+chr(9)+resi_county, resi_county
		' ComboBox 255, 255, 120, 45, county_list+chr(9)+resi_addr_county, resi_addr_county
		PushButton 290, 300, 95, 15, "Save Information", save_information_btn
	End If

	PushButton 325, 145, 50, 10, "CLEAR", clear_mail_addr_btn
	PushButton 205, 240, 35, 10, "CLEAR", clear_phone_one_btn
	PushButton 205, 260, 35, 10, "CLEAR", clear_phone_two_btn
	PushButton 205, 280, 35, 10, "CLEAR", clear_phone_three_btn
	' Text 10, 10, 360, 10, "Review the Address informaiton known with the client. If it needs updating, press this button to make changes:"
	GroupBox 10, 35, 375, 95, "Residence Address"
	Text 20, 55, 45, 10, "House/Street"
	Text 45, 75, 20, 10, "City"
	Text 185, 75, 20, 10, "State"
	Text 325, 75, 15, 10, "Zip"
	Text 20, 95, 100, 10, "Do you live on a Reservation?"
	Text 180, 95, 60, 10, "If yes, which one?"
	Text 30, 115, 90, 10, "Client Indicates Homeless:"
	Text 185, 115, 60, 10, "Living Situation?"
	GroupBox 10, 135, 375, 70, "Mailing Address"
	Text 20, 165, 45, 10, "House/Street"
	Text 45, 185, 20, 10, "City"
	Text 185, 185, 20, 10, "State"
	Text 325, 185, 15, 10, "Zip"
	GroupBox 10, 210, 235, 90, "Phone Number"
	Text 20, 225, 50, 10, "Number"
	Text 125, 225, 25, 10, "Type"
	Text 255, 225, 60, 10, "Date of Change:"
	Text 255, 245, 75, 10, "County of Residence:"
end function

function display_SHEL_information(update_shel, show_totals, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, update_information_btn, save_information_btn, const_memb_buttons, clear_all_btn, view_total_shel_btn, update_household_percent_button)

	Text 10, 10, 360, 10, "Review the Shelter informaiton known with the client. If it needs updating, press this button to make changes:"
	y_pos = 25
	For the_member = 0 to UBound(SHEL_ARRAY, 2)
		If the_member = selection Then
			Text 416, y_pos + 2, 60, 10, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member)
			y_pos = y_pos + 15
		Else
			PushButton 400, y_pos, 75, 13, "MEMBER " & SHEL_ARRAY(const_shel_member, the_member), SHEL_ARRAY(const_memb_buttons, the_member)
			y_pos = y_pos + 15
		End If
	Next
	' MsgBox "In DISPLAY" & vbCr & vbCr & "Show totals - " & show_totals
	If show_totals = True Then
		Text 415, 223, 65, 10, "TOTAL SHEL"

		If update_shel = True Then
			EditBox 105, 25, 165, 15, total_paid_to
			EditBox 125, 40, 20, 15, total_paid_by_household
			EditBox 125, 55, 20, 15, total_paid_by_others
			EditBox 105, 95, 45, 15, total_current_rent
			EditBox 105, 115, 45, 15, total_current_lot_rent
			EditBox 105, 135, 45, 15, total_current_mortgage
			EditBox 105, 155, 45, 15, total_current_insurance
			EditBox 105, 175, 45, 15, total_current_taxes
			EditBox 105, 195, 45, 15, total_current_room
			EditBox 105, 215, 45, 15, total_current_garage
			EditBox 105, 235, 45, 15, total_current_subsidy
			PushButton 400, 235, 75, 15, "Save Information", save_information_btn
		End If
		If update_shel = False Then
			Text 105, 30, 165, 10, total_paid_to
			Text 125, 45, 20, 10, total_paid_by_household
			Text 125, 60, 20, 10, total_paid_by_others
			Text 105, 100, 45, 10, total_current_rent
			Text 105, 120, 45, 10, total_current_lot_rent
			Text 105, 140, 45, 10, total_current_mortgage
			Text 105, 160, 45, 10, total_current_insurance
			Text 105, 180, 45, 10, total_current_taxes
			Text 105, 200, 45, 10, total_current_room
			Text 105, 220, 45, 10, total_current_garage
			Text 105, 240, 45, 10, total_current_subsidy
			PushButton 400, 235, 75, 15, "Update Information", update_information_btn
		End If
		Text 15, 30, 90, 10, "Housing Expense Paid to"
		Text 15, 45, 100, 10, "Expense Paid by Household"
		Text 145, 45, 50, 10, "% (percent)"
		PushButton 210, 41, 125, 13, "MANAGE HOUSEHOLD PERCENT", update_household_percent_button
		Text 15, 60, 100, 10, "Expense Paid by Someone Else"
		Text 145, 60, 50, 10, "% (percent)"
		Text 105, 75, 60, 20, "Current Total Amount"
		Text 80, 100, 20, 10, "Rent:"
	    Text 70, 120, 30, 10, "Lot Rent:"
	    Text 65, 140, 35, 10, "Mortgage:"
	    Text 65, 160, 40, 10, "Insurance:"
	    Text 75, 180, 25, 10, "Taxes:"
	    Text 75, 200, 25, 10, "Room:"
	    Text 75, 220, 30, 10, "Garage:"
	    Text 70, 240, 30, 10, "Subsidy:"

	Else
		PushButton 400, 220, 75, 15, "TOTAL SHEL", view_total_shel_btn

		If update_shel = True Then
			EditBox 105, 25, 165, 15, SHEL_ARRAY(const_paid_to, selection)
			DropListBox 165, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_hud_sub_yn, selection)
			DropListBox 310, 45, 40, 45, caf_answer_droplist, SHEL_ARRAY(const_shared_yn, selection)
			EditBox 105, 95, 45, 15, SHEL_ARRAY(const_rent_retro_amt, selection)
			DropListBox 155, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_retro_verif, selection)
			EditBox 255, 95, 45, 15, SHEL_ARRAY(const_rent_prosp_amt, selection)
			DropListBox 305, 95, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_rent_prosp_verif, selection)
			EditBox 105, 115, 45, 15, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
			DropListBox 155, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_retro_verif, selection)
			EditBox 255, 115, 45, 15, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
			DropListBox 305, 115, 85, 45, ""+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
			EditBox 105, 135, 45, 15, SHEL_ARRAY(const_mortgage_retro_amt, selection)
			DropListBox 155, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_retro_verif, selection)
			EditBox 255, 135, 45, 15, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
			DropListBox 305, 135, 85, 45, ""+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_mortgage_prosp_verif, selection)
			EditBox 105, 155, 45, 15, SHEL_ARRAY(const_insurance_retro_amt, selection)
			DropListBox 155, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_retro_verif, selection)
			EditBox 255, 155, 45, 15, SHEL_ARRAY(const_insurance_prosp_amt, selection)
			DropListBox 305, 155, 85, 45, ""+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_insurance_prosp_verif, selection)
			EditBox 105, 175, 45, 15, SHEL_ARRAY(const_tax_retro_amt, selection)
			DropListBox 155, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_retro_verif, selection)
			EditBox 255, 175, 45, 15, SHEL_ARRAY(const_tax_prosp_amt, selection)
			DropListBox 305, 175, 85, 45, ""+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_tax_prosp_verif, selection)
			EditBox 105, 195, 45, 15, SHEL_ARRAY(const_room_retro_amt, selection)
			DropListBox 155, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_retro_verif, selection)
			EditBox 255, 195, 45, 15, SHEL_ARRAY(const_room_prosp_amt, selection)
			DropListBox 305, 195, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_room_prosp_verif, selection)
			EditBox 105, 215, 45, 15, SHEL_ARRAY(const_garage_retro_amt, selection)
			DropListBox 155, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_retro_verif, selection)
			EditBox 255, 215, 45, 15, SHEL_ARRAY(const_garage_prosp_amt, selection)
			DropListBox 305, 215, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Chg, Neg Impact"+chr(9)+"PC - Chg, Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_garage_prosp_verif, selection)
			EditBox 105, 235, 45, 15, SHEL_ARRAY(const_subsidy_retro_amt, selection)
			DropListBox 155, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_retro_verif, selection)
			EditBox 255, 235, 45, 15, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
			DropListBox 305, 235, 85, 45, ""+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif", SHEL_ARRAY(const_subsidy_prosp_verif, selection)
			PushButton 400, 235, 75, 15, "Save Information", save_information_btn
		End If
		If update_shel = False Then
			Text 105, 30, 165, 10, SHEL_ARRAY(const_paid_to, selection)
			Text 165, 50, 40, 10, SHEL_ARRAY(const_hud_sub_yn, selection)
			Text 310, 50, 40, 10, SHEL_ARRAY(const_shared_yn, selection)
			Text 105, 100, 45, 10, SHEL_ARRAY(const_rent_retro_amt, selection)
			Text 160, 100, 70, 10, SHEL_ARRAY(const_rent_retro_verif, selection)
			Text 255, 100, 45, 10, SHEL_ARRAY(const_rent_prosp_amt, selection)
			Text 310, 100, 70, 10, SHEL_ARRAY(const_rent_prosp_verif, selection)
			Text 105, 120, 45, 10, SHEL_ARRAY(const_lot_rent_retro_amt, selection)
			Text 160, 120, 70, 10, SHEL_ARRAY(const_lot_rent_retro_verif, selection)
			Text 255, 120, 45, 10, SHEL_ARRAY(const_lot_rent_prosp_amt, selection)
			Text 310, 120, 70, 10, SHEL_ARRAY(const_lot_rent_prosp_verif, selection)
			Text 105, 140, 45, 10, SHEL_ARRAY(const_mortgage_retro_amt, selection)
			Text 160, 140, 70, 10, SHEL_ARRAY(const_mortgage_retro_verif, selection)
			Text 255, 140, 45, 10, SHEL_ARRAY(const_mortgage_prosp_amt, selection)
			Text 310, 140, 70, 10, SHEL_ARRAY(const_mortgage_prosp_verif, selection)
			Text 105, 160, 45, 10, SHEL_ARRAY(const_insurance_retro_amt, selection)
			Text 160, 160, 70, 10, SHEL_ARRAY(const_insurance_retro_verif, selection)
			Text 255, 160, 45, 10, SHEL_ARRAY(const_insurance_prosp_amt, selection)
			Text 310, 160, 70, 10, SHEL_ARRAY(const_insurance_prosp_verif, selection)
			Text 105, 180, 45, 10, SHEL_ARRAY(const_tax_retro_amt, selection)
			Text 160, 180, 70, 10, SHEL_ARRAY(const_tax_retro_verif, selection)
			Text 255, 180, 45, 10, SHEL_ARRAY(const_tax_prosp_amt, selection)
			Text 310, 180, 70, 10, SHEL_ARRAY(const_tax_prosp_verif, selection)
			Text 105, 200, 45, 10, SHEL_ARRAY(const_room_retro_amt, selection)
			Text 160, 200, 70, 10, SHEL_ARRAY(const_room_retro_verif, selection)
			Text 255, 200, 45, 10, SHEL_ARRAY(const_room_prosp_amt, selection)
			Text 310, 200, 70, 10, SHEL_ARRAY(const_room_prosp_verif, selection)
			Text 105, 220, 45, 10, SHEL_ARRAY(const_garage_retro_amt, selection)
			Text 160, 220, 70, 10, SHEL_ARRAY(const_garage_retro_verif, selection)
			Text 255, 220, 45, 10, SHEL_ARRAY(const_garage_prosp_amt, selection)
			Text 310, 220, 70, 10, SHEL_ARRAY(const_garage_prosp_verif, selection)
			Text 105, 240, 45, 10, SHEL_ARRAY(const_subsidy_retro_amt, selection)
			Text 160, 240, 70, 10, SHEL_ARRAY(const_subsidy_retro_verif, selection)
			Text 255, 240, 45, 10, SHEL_ARRAY(const_subsidy_prosp_amt, selection)
			Text 310, 240, 70, 10, SHEL_ARRAY(const_subsidy_prosp_verif, selection)
			PushButton 400, 235, 75, 15, "Update Information", update_information_btn
		End If

		PushButton 325, 30, 70, 13, "CLEAR ALL", clear_all_btn
	    Text 15, 30, 90, 10, "Housing Expense Paid to"
		Text 105, 50, 60, 10, "HUD Subsidized"
	    Text 225, 50, 85, 10, "Housing Expense Shared"
	    GroupBox 15, 65, 380, 190, "Housing Expense Amounts"
	    Text 105, 75, 65, 10, "Retrospective"
	    Text 255, 75, 65, 10, "Prospective"
	    Text 105, 85, 30, 10, "Amount"
	    Text 255, 85, 25, 10, "Amount"
	    Text 160, 85, 20, 10, "Verif"
	    Text 310, 85, 20, 10, "Verif"
		Text 80, 100, 20, 10, "Rent:"
	    Text 70, 120, 30, 10, "Lot Rent:"
	    Text 65, 140, 35, 10, "Mortgage:"
	    Text 65, 160, 40, 10, "Insurance:"
	    Text 75, 180, 25, 10, "Taxes:"
	    Text 75, 200, 25, 10, "Room:"
	    Text 75, 220, 30, 10, "Garage:"
	    Text 70, 240, 30, 10, "Subsidy:"

	End If



	'CAF Questions'
	' Text 20, 270, 125, 10, "Rent (include mobild home lot rental)"
    ' DropListBox 145, 265, 40, 45, "caf_answer_droplist", q14_rent_caf_answer
    ' EditBox 190, 265, 35, 15, q14_rent_caf_response
    ' Text 20, 285, 125, 10, "Mortgage/Contract for Deed Payment"
    ' DropListBox 145, 280, 40, 45, "caf_answer_droplist", q14_mort_caf_answer
    ' EditBox 190, 280, 35, 15, q14_mort_caf_response
    ' Text 20, 300, 125, 10, "Homeowner's Insurance"
    ' DropListBox 145, 295, 40, 45, "caf_answer_droplist", q14_ins_caf_answer
    ' EditBox 190, 295, 35, 15, q14_ins_caf_response
    ' Text 20, 315, 125, 10, "Real Estate Taxes"
    ' DropListBox 145, 310, 40, 45, "caf_answer_droplist", q14_tax_caf_answer
    ' EditBox 190, 310, 35, 15, q14_tax_caf_response
    ' Text 240, 270, 105, 10, "Rental or Secontion 8 Subsidy"
    ' DropListBox 345, 265, 40, 45, "caf_answer_droplist", q14_subs_caf_answer
    ' EditBox 390, 265, 35, 15, q14_subs_caf_response
    ' Text 240, 285, 100, 10, "Association Fees"
    ' DropListBox 345, 280, 40, 45, "caf_answer_droplist", q14_fees_caf_answer
    ' EditBox 390, 280, 35, 15, q14_fees_caf_response
    ' Text 240, 300, 95, 10, "Room and/or Board"
    ' DropListBox 345, 295, 40, 45, "caf_answer_droplist", q14_room_caf_answer
    ' EditBox 390, 295, 35, 15, q14_room_caf_response
    ' Text 240, 315, 105, 20, "CONFIM - Do you get help paying rent?"
    ' DropListBox 345, 310, 40, 45, "caf_answer_droplist", q14_confirm_subsidy
    ' EditBox 390, 310, 35, 15, q14_confirm_subsidy_amount
end function

function display_HEST_information(update_hest, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

	If update_hest = False Then
		Text 75, 30, 145, 10, all_persons_paying
	    Text 75, 50, 50, 10, choice_date
	    Text 125, 70, 50, 10, actual_initial_exp
	    Text 70, 125, 40, 10, retro_heat_ac_yn
	    Text 115, 125, 20, 10, retro_heat_ac_units
	    Text 150, 125, 45, 10, retro_heat_ac_amt
	    Text 240, 125, 40, 10, prosp_heat_ac_yn
	    Text 285, 125, 20, 10, prosp_heat_ac_units
	    Text 320, 125, 45, 10, prosp_heat_ac_amt
	    Text 70, 145, 40, 10, retro_electric_yn
	    Text 115, 145, 20, 10, retro_electric_units
	    Text 150, 145, 45, 10, retro_electric_amt
	    Text 240, 145, 40, 10, prosp_electric_yn
	    Text 285, 145, 20, 10, prosp_electric_units
	    Text 320, 145, 45, 10, prosp_electric_amt
	    Text 70, 165, 40, 10, retro_phone_yn
	    Text 115, 165, 20, 10, retro_phone_units
	    Text 150, 165, 45, 10, retro_phone_amt
	    Text 240, 165, 40, 10, prosp_phone_yn
	    Text 285, 165, 20, 10, prosp_phone_units
	    Text 320, 165, 45, 10, prosp_phone_amt
		Text 55, 185, 150, 10, "Total Counted Utility Expense: $" & total_utility_expense

		PushButton 290, 185, 95, 15, "Update Information", update_information_btn
	End If
	If update_hest = True Then
		EditBox 75, 25, 145, 15, all_persons_paying
	    EditBox 75, 45, 50, 15, choice_date
	    EditBox 125, 65, 50, 15, actual_initial_exp
	    DropListBox 65, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_heat_ac_yn
	    ' EditBox 110, 120, 20, 15, retro_heat_ac_units
	    ' EditBox 150, 120, 45, 15, retro_heat_ac_amt
	    DropListBox 235, 120, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_heat_ac_yn
	    ' EditBox 280, 120, 20, 15, prosp_heat_ac_units
	    ' EditBox 320, 120, 45, 15, prosp_heat_ac_amt
	    DropListBox 65, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_electric_yn
	    ' EditBox 110, 140, 20, 15, retro_electric_units
	    ' EditBox 150, 140, 45, 15, retro_electric_amt
	    DropListBox 235, 140, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_electric_yn
	    ' EditBox 280, 140, 20, 15, prosp_electric_units
	    ' EditBox 320, 140, 45, 15, prosp_electric_amt
	    DropListBox 65, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", retro_phone_yn
	    ' EditBox 110, 160, 20, 15, retro_phone_units
	    ' EditBox 150, 160, 45, 15, retro_phone_amt
	    DropListBox 235, 160, 30, 45, ""+chr(9)+"Y"+chr(9)+"N", prosp_phone_yn
	    ' EditBox 280, 160, 20, 15, prosp_phone_units
	    ' EditBox 320, 160, 45, 15, prosp_phone_amt
		' ComboBox 255, 255, 120, 45, county_list+chr(9)+resi_addr_county, resi_addr_county
		PushButton 290, 185, 95, 15, "Save Information", save_information_btn
	End If


	Text 10, 10, 360, 10, "Review the Utility Information"
    Text 15, 30, 60, 10, "Persons Paying:"
    Text 15, 50, 55, 10, "FS Choice Date:"
    Text 15, 70, 110, 10, "Actual Expense In Initial Month: $ "
    Text 20, 125, 30, 10, "Heat/Air:"
    Text 20, 145, 30, 10, "Electric:"
    Text 25, 165, 25, 10, "Phone:"
    GroupBox 55, 85, 150, 95, "Retrospective"
    Text 65, 105, 20, 10, "(Y/N)"
    Text 110, 100, 20, 20, "#/FS Units"
    Text 150, 105, 30, 10, "Amount"
    GroupBox 225, 85, 150, 95, "Prospective"
    Text 235, 105, 20, 10, "(Y/N)"
    Text 280, 100, 20, 20, "#/FS Units"
    Text 320, 105, 25, 10, "Amount"

	' GroupBox 20, 150, 455, grp_len, "Already Known Shelter Expenses - Added or listed in MAXIS"
	' ' Text 30, 165, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
	' ' Text 30, 180, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
	' PushButton 350, y_pos, 125, 10, "Update Shelter Expense Information", update_shel_btn
	' y_pos = y_pos + 15
	' Text 5, y_pos, 310, 10, "^^4 - Enter the answers listed on the actual CAF form for Q15 into the 'Answer on the CAF' field."
	' Text 20, y_pos + 10, 295, 10, "Q. 15. Does your household have the following utility expenses any time during the year?"
	' y_pos = y_pos + 30
	' Text 20, y_pos, 85, 10, "Heating/Air Conditioning"
	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_ac_caf_answer
	' Text 180, y_pos, 85, 10, "Electricity"
	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_answer
	' Text 345, y_pos, 85, 10, "Cooking Fuel"
	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_answer
	' y_pos = y_pos + 15
	' Text 20, y_pos, 85, 10, "Water and Sewer"
	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_answer
	' Text 180, y_pos, 85, 10, "Garbage Removal"
	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_answer
	' Text 345, y_pos, 85, 10, "Phone/Cell Phone"
	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_answer
	' y_pos = y_pos + 15
	' Text 75, y_pos, 355, 10, "Did anyone in the household receive Energy Assistance (LIHEAP) of more than $20 in the past 12 months?"
	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_liheap_caf_answer
	' y_pos = y_pos + 15
	'
	' Text 5, y_pos, 270, 10, "^^5 - ASK - 'Does anyone in the household pay ...'  RECORD the verbal responses"
	' y_pos = y_pos + 20
	' Text 20, y_pos, 85, 10, "Heating"
	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_caf_response
	' Text 180, y_pos, 85, 10, "Electricity"
	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_response
	' Text 345, y_pos, 85, 10, "Cooking Fuel"
	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_response
	' y_pos = y_pos + 15
	' Text 20, y_pos, 85, 10, "Air Conditioning"
	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ac_caf_response
	' Text 180, y_pos, 85, 10, "Garbage Removal"
	' DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_response
	' Text 345, y_pos, 85, 10, "Phone/Cell Phone"
	' DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_response
	' y_pos = y_pos + 15
	' Text 20, y_pos, 85, 10, "Water and Sewer"
	' DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_response
	' Text 170, y_pos + 5, 265, 10, "Did your household receive any help in paying for your energy or power bills?"
	' DropListBox 435, y_pos, 40, 45, caf_answer_droplist, q15_liheap_caf_response
	' y_pos = y_pos + 15
	' PushButton 20, y_pos, 130, 10, "Utilities are Complicated", utility_detail_btn
end function

function get_state_name_from_state_code(state_code, state_name, include_state_code)
    If state_code = "NB" Then state_name = "MN Newborn"							'This is the list of all the states connected to the code.
    If state_code = "FC" Then state_name = "Foreign Country"
    If state_code = "UN" Then state_name = "Unknown"
    If state_code = "AL" Then state_name = "Alabama"
    If state_code = "AK" Then state_name = "Alaska"
    If state_code = "AZ" Then state_name = "Arizona"
    If state_code = "AR" Then state_name = "Arkansas"
    If state_code = "CA" Then state_name = "California"
    If state_code = "CO" Then state_name = "Colorado"
    If state_code = "CT" Then state_name = "Connecticut"
    If state_code = "DE" Then state_name = "Delaware"
    If state_code = "DC" Then state_name = "District Of Columbia"
    If state_code = "FL" Then state_name = "Florida"
    If state_code = "GA" Then state_name = "Georgia"
    If state_code = "HI" Then state_name = "Hawaii"
    If state_code = "ID" Then state_name = "Idaho"
    If state_code = "IL" Then state_name = "Illnois"
    If state_code = "IN" Then state_name = "Indiana"
    If state_code = "IA" Then state_name = "Iowa"
    If state_code = "KS" Then state_name = "Kansas"
    If state_code = "KY" Then state_name = "Kentucky"
    If state_code = "LA" Then state_name = "Louisiana"
    If state_code = "ME" Then state_name = "Maine"
    If state_code = "MD" Then state_name = "Maryland"
    If state_code = "MA" Then state_name = "Massachusetts"
    If state_code = "MI" Then state_name = "Michigan"
	If state_code = "MN" Then state_name = "Minnesota"
    If state_code = "MS" Then state_name = "Mississippi"
    If state_code = "MO" Then state_name = "Missouri"
    If state_code = "MT" Then state_name = "Montana"
    If state_code = "NE" Then state_name = "Nebraska"
    If state_code = "NV" Then state_name = "Nevada"
    If state_code = "NH" Then state_name = "New Hampshire"
    If state_code = "NJ" Then state_name = "New Jersey"
    If state_code = "NM" Then state_name = "New Mexico"
    If state_code = "NY" Then state_name = "New York"
    If state_code = "NC" Then state_name = "North Carolina"
    If state_code = "ND" Then state_name = "North Dakota"
    If state_code = "OH" Then state_name = "Ohio"
    If state_code = "OK" Then state_name = "Oklahoma"
    If state_code = "OR" Then state_name = "Oregon"
    If state_code = "PA" Then state_name = "Pennsylvania"
    If state_code = "RI" Then state_name = "Rhode Island"
    If state_code = "SC" Then state_name = "South Carolina"
    If state_code = "SD" Then state_name = "South Dakota"
    If state_code = "TN" Then state_name = "Tennessee"
    If state_code = "TX" Then state_name = "Texas"
    If state_code = "UT" Then state_name = "Utah"
    If state_code = "VT" Then state_name = "Vermont"
    If state_code = "VA" Then state_name = "Virginia"
    If state_code = "WA" Then state_name = "Washington"
    If state_code = "WV" Then state_name = "West Virginia"
    If state_code = "WI" Then state_name = "Wisconsin"
    If state_code = "WY" Then state_name = "Wyoming"
    If state_code = "PR" Then state_name = "Puerto Rico"
    If state_code = "VI" Then state_name = "Virgin Islands"

    If include_state_code = TRUE Then state_name = state_code & " " & state_name	'This adds the code to the state name if seelected
end function

function navigate_ADDR_buttons(update_addr, err_var, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn, mail_street_full, mail_city, mail_state, mail_zip, phone_one, phone_two, phone_three, type_one, type_two, type_three)
	If ButtonPressed = update_information_btn Then
		update_addr = TRUE
		' update_attempted = True
	ElseIf ButtonPressed = save_information_btn Then
		update_addr = FALSE
	Else
		update_addr = FALSE
	End If

	If ButtonPressed = clear_mail_addr_btn Then
		mail_street_full = ""
		mail_city = ""
		mail_state = ""
		mail_zip = ""
	End If
	If ButtonPressed = clear_phone_one_btn Then
		phone_one = ""
		type_one = "Select One..."
	End If
	If ButtonPressed = clear_phone_two_btn Then
		phone_two = ""
		type_two = "Select One..."
	End If
	If ButtonPressed = clear_phone_three_btn Then
		phone_three = ""
		type_three = "Select One..."
	End If
end function

' function navigate_SHEL_buttons(update_shel, err_var, update_attempted, update_information_btn, save_information_btn, SHEL_ARRAY, const_memb_buttons, const_shel_exists, const_attempt_update, selection)

function navigate_SHEL_buttons(update_shel, show_totals, err_var, SHEL_ARRAY, selection, const_shel_member, const_shel_exists, const_hud_sub_yn, const_shared_yn, const_paid_to, const_rent_retro_amt, const_rent_retro_verif, const_rent_prosp_amt, const_rent_prosp_verif, const_lot_rent_retro_amt, const_lot_rent_retro_verif, const_lot_rent_prosp_amt, const_lot_rent_prosp_verif, const_mortgage_retro_amt, const_mortgage_retro_verif, const_mortgage_prosp_amt, const_mortgage_prosp_verif, const_insurance_retro_amt, const_insurance_retro_verif, const_insurance_prosp_amt, const_insurance_prosp_verif, const_tax_retro_amt, const_tax_retro_verif, const_tax_prosp_amt, const_tax_prosp_verif, const_room_retro_amt, const_room_retro_verif, const_room_prosp_amt, const_room_prosp_verif, const_garage_retro_amt, const_garage_retro_verif, const_garage_prosp_amt, const_garage_prosp_verif, const_subsidy_retro_amt, const_subsidy_retro_verif, const_subsidy_prosp_amt, const_subsidy_prosp_verif, update_information_btn, save_information_btn, const_memb_buttons, const_attempt_update, clear_all_btn, view_total_shel_btn)

	If ButtonPressed = update_information_btn Then
		update_shel = TRUE
		update_attempted = True
		' MsgBox "In UPDATE button" & vbCr & vbCr & "Show totals - " & show_totals
	ElseIf ButtonPressed = save_information_btn Then
		update_shel = FALSE
	Else
		update_shel = FALSE
	End If

	If selection <> "" Then
		'REVIEWING THE INFORMATION IN THE ARRAY TO DETERMINE IF IT IS BLANK
		all_shel_details_blank = True

		If Trim(SHEL_ARRAY(const_paid_to, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_hud_sub_yn, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_shared_yn, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_lot_rent_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_mortgage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_insurance_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_tax_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_room_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_garage_prosp_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_retro_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_retro_verif, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_prosp_amt, selection)) <> "" Then all_shel_details_blank = False
		If Trim(SHEL_ARRAY(const_subsidy_prosp_verif, selection)) <> "" Then all_shel_details_blank = False

		If all_shel_details_blank = True Then SHEL_ARRAY(const_shel_exists, selection) = False

		If ButtonPressed = clear_all_btn Then
			SHEL_ARRAY(const_paid_to, selection) = ""
			SHEL_ARRAY(const_hud_sub_yn, selection) = ""
			SHEL_ARRAY(const_shared_yn, selection) = ""
			SHEL_ARRAY(const_rent_retro_amt, selection) = ""
			SHEL_ARRAY(const_rent_retro_verif, selection) = ""
			SHEL_ARRAY(const_rent_prosp_amt, selection) = ""
			SHEL_ARRAY(const_rent_prosp_verif, selection) = ""
			SHEL_ARRAY(const_lot_rent_retro_amt, selection) = ""
			SHEL_ARRAY(const_lot_rent_retro_verif, selection) = ""
			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = ""
			SHEL_ARRAY(const_lot_rent_prosp_verif, selection) = ""
			SHEL_ARRAY(const_mortgage_retro_amt, selection) = ""
			SHEL_ARRAY(const_mortgage_retro_verif, selection) = ""
			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = ""
			SHEL_ARRAY(const_mortgage_prosp_verif, selection) = ""
			SHEL_ARRAY(const_insurance_retro_amt, selection) = ""
			SHEL_ARRAY(const_insurance_retro_verif, selection) = ""
			SHEL_ARRAY(const_insurance_prosp_amt, selection) = ""
			SHEL_ARRAY(const_insurance_prosp_verif, selection) = ""
			SHEL_ARRAY(const_tax_retro_amt, selection) = ""
			SHEL_ARRAY(const_tax_retro_verif, selection) = ""
			SHEL_ARRAY(const_tax_prosp_amt, selection) = ""
			SHEL_ARRAY(const_tax_prosp_verif, selection) = ""
			SHEL_ARRAY(const_room_retro_amt, selection) = ""
			SHEL_ARRAY(const_room_retro_verif, selection) = ""
			SHEL_ARRAY(const_room_prosp_amt, selection) = ""
			SHEL_ARRAY(const_room_prosp_verif, selection) = ""
			SHEL_ARRAY(const_garage_retro_amt, selection) = ""
			SHEL_ARRAY(const_garage_retro_verif, selection) = ""
			SHEL_ARRAY(const_garage_prosp_amt, selection) = ""
			SHEL_ARRAY(const_garage_prosp_verif, selection) = ""
			SHEL_ARRAY(const_subsidy_retro_amt, selection) = ""
			SHEL_ARRAY(const_subsidy_retro_verif, selection) = ""
			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = ""
			SHEL_ARRAY(const_subsidy_prosp_verif, selection) = ""
			SHEL_ARRAY(const_shel_exists, selection) = False
		End If
	End If

	For memb_btn = 0 to UBound(SHEL_ARRAY, 2)
		If ButtonPressed = SHEL_ARRAY(const_memb_buttons, memb_btn) Then
			selection = memb_btn
			show_totals = False
		End If
	Next
	If selection <> "" Then
		If SHEL_ARRAY(const_shel_exists, selection) = False Then update_shel = True
		If update_shel = True Then
			SHEL_ARRAY(const_attempt_update, selection) = True
			update_attempted = True

			SHEL_ARRAY(const_rent_prosp_amt, selection) = SHEL_ARRAY(const_rent_prosp_amt, selection) & ""
			SHEL_ARRAY(const_lot_rent_prosp_amt, selection) = SHEL_ARRAY(const_lot_rent_prosp_amt, selection) & ""
			SHEL_ARRAY(const_mortgage_prosp_amt, selection) = SHEL_ARRAY(const_mortgage_prosp_amt, selection) & ""
			SHEL_ARRAY(const_insurance_prosp_amt, selection) = SHEL_ARRAY(const_insurance_prosp_amt, selection) & ""
			SHEL_ARRAY(const_tax_prosp_amt, selection) = SHEL_ARRAY(const_tax_prosp_amt, selection) & ""
			SHEL_ARRAY(const_room_prosp_amt, selection) = SHEL_ARRAY(const_room_prosp_amt, selection) & ""
			SHEL_ARRAY(const_garage_prosp_amt, selection) = SHEL_ARRAY(const_garage_prosp_amt, selection) & ""
			SHEL_ARRAY(const_subsidy_prosp_amt, selection) = SHEL_ARRAY(const_subsidy_prosp_amt, selection) & ""
		End If
	End If
	If ButtonPressed = view_total_shel_btn Then
		show_totals = True
		selection = ""
	End If
	If show_totals = True and update_shel = True Then
		total_paid_by_household = total_paid_by_household & ""
		total_paid_by_others = total_paid_by_others & ""
		total_current_rent = total_current_rent & ""
		total_current_lot_rent = total_current_lot_rent & ""
		total_current_mortgage = total_current_mortgage & ""
		total_current_insurance = total_current_insurance & ""
		total_current_taxes = total_current_taxes & ""
		total_current_room = total_current_room & ""
		total_current_garage = total_current_garage & ""
		total_current_subsidy = total_current_subsidy & ""
	End If
	' MsgBox "End NAVIGATE" & vbCr & vbCr & "Show totals - " & show_totals
end function

function navigate_HEST_buttons(update_hest, err_var, update_attempted, update_information_btn, save_information_btn, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
	If ButtonPressed = update_information_btn Then
		update_hest = TRUE
		update_attempted = True

		retro_heat_ac_amt = retro_heat_ac_amt & ""
		retro_electric_amt = retro_electric_amt & ""
		retro_phone_amt = retro_phone_amt & ""
		prosp_heat_ac_amt = prosp_heat_ac_amt & ""
		prosp_electric_amt = prosp_electric_amt & ""
		prosp_phone_amt = prosp_phone_amt & ""

	ElseIf ButtonPressed = save_information_btn Then
		update_hest = FALSE

		retro_heat_ac_amt = 0
		retro_heat_ac_units = ""
		retro_electric_amt = 0
		retro_electric_units = ""
		retro_phone_amt = 0
		retro_phone_units = ""
		prosp_heat_ac_amt = 0
		prosp_heat_ac_units = ""
		prosp_electric_amt = 0
		prosp_electric_units = ""
		prosp_phone_amt = 0
		prosp_phone_units = ""

		If retro_heat_ac_yn = "Y" Then
			retro_heat_ac_amt = 496
			retro_heat_ac_units = "01"
		End If
		If retro_electric_yn = "Y" Then
			retro_electric_amt = 154
			retro_electric_units = "01"
		End If
		If retro_phone_yn = "Y" Then
			retro_phone_amt = 56
			retro_phone_units = "01"
		End If
		If prosp_heat_ac_yn = "Y" Then
			prosp_heat_ac_amt = 496
			prosp_heat_ac_units = "01"
		End If
		If prosp_electric_yn = "Y" Then
			prosp_electric_amt = 154
			prosp_electric_units = "01"
		End If
		If prosp_phone_yn = "Y" Then
			prosp_phone_amt = 56
			prosp_phone_units = "01"
		End If

		total_utility_expense = 0
		If prosp_heat_ac_yn = "Y" Then
			total_utility_expense =  prosp_heat_ac_amt
		ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
			total_utility_expense =  prosp_electric_amt + prosp_phone_amt
		ElseIf prosp_electric_yn = "Y" Then
			total_utility_expense =  prosp_electric_amt
		Elseif prosp_phone_yn = "Y" Then
			total_utility_expense =  prosp_phone_amt
		End If

	Else
		update_hest = FALSE
	End If
end function

function navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, total_current_taxes, total_current_lot_rent, total_current_room, total_current_mortgage, total_current_garage, total_current_insurance, total_current_subsidy, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, update_shel_button, housing_change_continue_btn, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, addr_update_needed, shel_update_needed, hest_update_needed)

	' view_addr_update_dlg
	' view_shel_update_dlg
	' view_shel_details_dlg
	If housing_questions_step = 3 Then

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
	If err_msg = "" Then
		If ButtonPressed = housing_change_continue_btn Then
			housing_questions_step = housing_questions_step + 1

			If housing_questions_step = 2 and view_addr_update_dlg = False Then housing_questions_step = housing_questions_step + 1
			If housing_questions_step = 3 and view_shel_update_dlg = False Then housing_questions_step = housing_questions_step + 1
			If housing_questions_step = 4 and view_shel_details_dlg = False Then housing_questions_step = housing_questions_step + 1

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
		If ButtonPressed = housing_change_shel_details_btn Then housing_questions_step = 4

		If housing_questions_step = 3 Then

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
end function

function read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, rent_amt, rent_verif, lot_rent_amt, lot_rent_verif, mortgage_amt, mortgage_verif, insurance_amt, insurance_verif, taxes_amt, taxes_verif, room_amt, room_verif, garage_amt, garage_verif, subsidy_amt, subsidy_verif, original_information)

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
			client_list = client_list & ref_nbr & "|"
		End If
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_list = TRIM(client_list)
	If right(client_list, 1) = "|" Then client_list = left(client_list, len(client_list) - 1)
	shel_ref_numbers_array = split(client_list, "|")

	rent_amt = 0
	rent_verif = ""
	lot_rent_amt = 0
	lot_rent_verif = ""
	mortgage_amt = 0
	mortgage_verif = ""
	insurance_amt = 0
	insurance_verif = ""
	taxes_amt = 0
	taxes_verif = ""
	room_amt = 0
	room_verif = ""
	garage_amt = 0
	garage_verif = ""
	subsidy_amt = 0
	subsidy_verif = ""

	Call navigate_to_MAXIS_screen("STAT", "SHEL")

	For each memb_ref_number in shel_ref_numbers_array
		EMWriteScreen memb_ref_number, 20, 76
		transmit

		EMReadScreen shel_version, 1, 2, 78
		If shel_version = "1" Then
			ref_numbers_with_panel = ref_numbers_with_panel & "~" & memb_ref_number

		    EMReadScreen panel_paid_to,               25, 7, 50
		    panel_paid_to = replace(panel_paid_to, "_", "")
			If paid_to = "" Then
				paid_to = panel_paid_to
			ElseIf paid_to <> panel_paid_to Then
				paid_to = "Multiple"
			End If

		    EMReadScreen rent_prosp_amt,        8, 11, 56
		    EMReadScreen rent_prosp_verif,      2, 11, 67

		    rent_prosp_amt = replace(rent_prosp_amt, "_", "")
		    rent_prosp_amt = trim(rent_prosp_amt)
			If rent_prosp_amt = "" Then rent_prosp_amt = 0
			rent_prosp_amt = rent_prosp_amt * 1
			rent_amt = rent_amt + rent_prosp_amt

		    If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
		    If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
		    If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
		    If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Doc"
		    If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg, Neg Impact"
		    If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg, Pos Impact"
		    If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Verif"
			If rent_prosp_verif = "?_" Then rent_prosp_verif = "? - Delayed Verif"
		    If rent_prosp_verif = "__" Then rent_prosp_verif = ""
			If rent_verif = "" Then
				rent_verif = rent_prosp_verif
			ElseIf rent_verif <> rent_prosp_verif Then
				rent_verif = "Multiple"
			End If

		    EMReadScreen lot_rent_prosp_amt,    8, 12, 56
		    EMReadScreen lot_rent_prosp_verif,  2, 12, 67

		    lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
		    lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
			If lot_rent_prosp_amt = "" Then lot_rent_prosp_amt = 0
			lot_rent_prosp_amt = lot_rent_prosp_amt * 1
			lot_rent_amt = lot_rent_amt + lot_rent_prosp_amt
		    If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
		    If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
		    If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
		    If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Doc"
		    If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg, Neg Impact"
		    If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg, Pos Impact"
		    If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Verif"
			If lot_rent_prosp_verif = "?_" Then lot_rent_prosp_verif = "? - Delayed Verif"
		    If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""
			If lot_rent_verif = "" Then
				lot_rent_verif = lot_rent_prosp_verif
			ElseIf lot_rent_verif <> lot_rent_prosp_verif Then
				lot_rent_verif = "Multiple"
			End If

		    EMReadScreen mortgage_prosp_amt,    8, 13, 56
		    EMReadScreen mortgage_prosp_verif,  2, 13, 67

		    mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
		    mortgage_prosp_amt = trim(mortgage_prosp_amt)
			If mortgage_prosp_amt = "" Then mortgage_prosp_amt = 0
			mortgage_prosp_amt = mortgage_prosp_amt * 1
			mortgage_amt = mortgage_amt + mortgage_prosp_amt
		    If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
		    If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
		    If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Doc"
		    If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg, Neg Impact"
		    If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg, Pos Impact"
		    If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Verif"
			If mortgage_prosp_verif = "?_" Then mortgage_prosp_verif = "? - Delayed Verif"
		    If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""
			If mortgage_verif = "" Then
				mortgage_verif = mortgage_prosp_verif
			ElseIf mortgage_verif <> mortgage_prosp_verif Then
				mortgage_verif = "Multiple"
			End If

		    EMReadScreen insurance_prosp_amt,   8, 14, 56
		    EMReadScreen insurance_prosp_verif, 2, 14, 67

		    insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
		    insurance_prosp_amt = trim(insurance_prosp_amt)
			If insurance_prosp_amt = "" Then insurance_prosp_amt = 0
			insurance_prosp_amt = insurance_prosp_amt * 1
			insurance_amt = insurance_amt + insurance_prosp_amt
		    If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
		    If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Doc"
		    If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg, Neg Impact"
		    If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg, Pos Impact"
		    If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Verif"
			If insurance_prosp_verif = "?_" Then insurance_prosp_verif = "? - Delayed Verif"
		    If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""
			If insurance_verif = "" Then
				insurance_verif = insurance_prosp_verif
			ElseIf insurance_verif <> insurance_prosp_verif Then
				insurance_verif = "Multiple"
			End If

		    EMReadScreen tax_prosp_amt,         8, 15, 56
		    EMReadScreen tax_prosp_verif,       2, 15, 67

		    tax_prosp_amt = replace(tax_prosp_amt, "_", "")
		    tax_prosp_amt = trim(tax_prosp_amt)
			If tax_prosp_amt = "" Then tax_prosp_amt = 0
			tax_prosp_amt = tax_prosp_amt * 1
			taxes_amt = taxes_amt + tax_prosp_amt
		    If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
		    If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Doc"
		    If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg, Neg Impact"
		    If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg, Pos Impact"
		    If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Verif"
			If tax_prosp_verif = "?_" Then tax_prosp_verif = "? - Delayed Verif"
		    If tax_prosp_verif = "__" Then tax_prosp_verif = ""
			If taxes_verif = "" Then
				taxes_verif = tax_prosp_verif
			ElseIf taxes_verif <> tax_prosp_verif Then
				taxes_verif = "Multiple"
			End If

		    EMReadScreen room_prosp_amt,        8, 16, 56
		    EMReadScreen room_prosp_verif,      2, 16, 67

		    room_prosp_amt = replace(room_prosp_amt, "_", "")
		    room_prosp_amt = trim(room_prosp_amt)
			If room_prosp_amt = "" Then room_prosp_amt = 0
			room_prosp_amt = room_prosp_amt * 1
			room_amt = room_amt + room_prosp_amt
		    If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
		    If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
		    If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
		    If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Doc"
		    If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg, Neg Impact"
		    If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg, Pos Impact"
		    If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Verif"
			If room_prosp_verif = "?_" Then room_prosp_verif = "? - Delayed Verif"
		    If room_prosp_verif = "__" Then room_prosp_verif = ""
			If room_verif = "" Then
				room_verif = room_prosp_verif
			ElseIf room_verif <> room_prosp_verif Then
				room_verif = "Multiple"
			End If

		    EMReadScreen garage_prosp_amt,      8, 17, 56
		    EMReadScreen garage_prosp_verif,    2, 17, 67

		    garage_prosp_amt = replace(garage_prosp_amt, "_", "")
		    garage_prosp_amt = trim(garage_prosp_amt)
			If garage_prosp_amt = "" Then garage_prosp_amt = 0
			garage_prosp_amt = garage_prosp_amt * 1
			garage_amt = garage_amt + garage_prosp_amt
		    If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
		    If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
		    If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
		    If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Doc"
		    If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg, Neg Impact"
		    If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg, Pos Impact"
		    If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Verif"
			If garage_prosp_verif = "?_" Then garage_prosp_verif = "? - Delayed Verif"
		    If garage_prosp_verif = "__" Then garage_prosp_verif = ""
			If garage_verif = "" Then
				garage_verif = garage_prosp_verif
			ElseIf garage_verif <> garage_prosp_verif Then
				garage_verif = "Multiple"
			End If

		    EMReadScreen subsidy_prosp_amt,     8, 18, 56
		    EMReadScreen subsidy_prosp_verif,   2, 18, 67

		    subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
		    subsidy_prosp_amt = trim(subsidy_prosp_amt)
			If subsidy_prosp_amt = "" Then subsidy_prosp_amt = 0
			subsidy_prosp_amt = subsidy_prosp_amt * 1
			subsidy_amt = subsidy_amt + subsidy_prosp_amt
		    If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
		    If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
		    If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Doc"
		    If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Verif"
			If subsidy_prosp_verif = "?_" Then subsidy_prosp_verif = "? - Delayed Verif"
		    If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""
			If subsidy_verif = "" Then
				subsidy_verif = subsidy_prosp_verif
			ElseIf subsidy_verif <> subsidy_prosp_verif Then
				subsidy_verif = "Multiple"
			End If
		End If
	Next

	original_information = rent_amt&"|"&rent_verif&"|"&lot_rent_amt&"|"&lot_rent_verif&"|"&mortgage_amt&"|"&mortgage_verif&"|"&insurance_amt&"|"&insurance_verif&"|"&taxes_amt&"|"&taxes_verif&"|"&room_amt&"|"&room_verif&"|"&garage_amt&"|"&garage_verif&"|"&subsidy_amt&"|"&subsidy_verif
end function

function split_phone_number_into_parts(phone_variable, phone_left, phone_mid, phone_right)
'This function is to take the information provided as a phone number and split it up into the 3 parts
    phone_variable = trim(phone_variable)
    If phone_variable <> "" Then
        phone_variable = replace(phone_variable, "(", "")						'formatting the phone variable to get rid of symbols and spaces
        phone_variable = replace(phone_variable, ")", "")
        phone_variable = replace(phone_variable, "-", "")
        phone_variable = replace(phone_variable, " ", "")
        phone_variable = trim(phone_variable)
        phone_left = left(phone_variable, 3)									'reading the certain sections of the variable for each part.
        phone_mid = mid(phone_variable, 4, 3)
        phone_right = right(phone_variable, 4)
        phone_variable = "(" & phone_left & ")" & phone_mid & "-" & phone_right
    End If
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
const original_panel_info_const		= 39
const shel_entered_notes_const		= 40

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
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 196, 60, "Dialog"
  DropListBox 15, 15, 170, 45, " "+chr(9)+"Application/Renewal"+chr(9)+"Change", select_option
  ButtonGroup ButtonPressed
    OkButton 135, 35, 50, 15
EndDialog

dialog Dialog1

If select_option = "Application/Renewal" Then

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
			ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member) = "||||||||||||||||||||||||||||||||||"
		End If
		members_counter = members_counter + 1
	Next

	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_addr_panel_info, addr_update_attempted)
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

	addr_update_attempted = False
	shel_update_attempted = False
	hest_update_attempted = False

	Do
		err_msg = ""

		BeginDialog Dialog1, 0, 0, 555, 385, "Housing Expense Detail"

		  ButtonGroup ButtonPressed

		  	If page_to_display = ADDR_dlg_page Then
				Text 506, 12, 60, 10, "ADDR"
				Call display_ADDR_information(update_addr, notes_on_address, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, reservation_name, addr_living_sit, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, address_change_date, update_information_btn, save_information_btn, clear_mail_addr_btn, clear_phone_one_btn, clear_phone_two_btn, clear_phone_three_btn)
			End If

			If page_to_display = SHEL_dlg_page Then
				Text 506, 27, 60, 10, "SHEL"

				Call display_SHEL_information(update_shel, show_totals, ALL_SHEL_PANELS_ARRAY, member_selection, shel_ref_number_const, shel_exists_const, hud_sub_yn_const, shared_yn_const, paid_to_const, rent_retro_amt_const, rent_retro_verif_const, rent_prosp_amt_const, rent_prosp_verif_const, lot_rent_retro_amt_const, lot_rent_retro_verif_const, lot_rent_prosp_amt_const, lot_rent_prosp_verif_const, mortgage_retro_amt_const, mortgage_retro_verif_const, mortgage_prosp_amt_const, mortgage_prosp_verif_const, insurance_retro_amt_const, insurance_retro_verif_const, insurance_prosp_amt_const, insurance_prosp_verif_const, tax_retro_amt_const, tax_retro_verif_const, tax_prosp_amt_const, tax_prosp_verif_const, room_retro_amt_const, room_retro_verif_const, room_prosp_amt_const, room_prosp_verif_const, garage_retro_amt_const, garage_retro_verif_const, garage_prosp_amt_const, garage_prosp_verif_const, subsidy_retro_amt_const, subsidy_retro_verif_const, subsidy_prosp_amt_const, subsidy_prosp_verif_const, update_information_btn, save_information_btn, memb_btn_const, clear_all_btn, view_total_shel_btn, update_household_percent_button)
			End If

			If page_to_display = HEST_dlg_page Then
				Text 507, 42, 60, 10, "HEST"
				Call display_HEST_information(update_hest, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

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
		If page_to_display = SHEL_dlg_page Then Call navigate_SHEL_buttons(update_shel, show_totals, err_var, ALL_SHEL_PANELS_ARRAY, member_selection, shel_ref_number_const, shel_exists_const, hud_sub_yn_const, shared_yn_const, paid_to_const, rent_retro_amt_const, rent_retro_verif_const, rent_prosp_amt_const, rent_prosp_verif_const, lot_rent_retro_amt_const, lot_rent_retro_verif_const, lot_rent_prosp_amt_const, lot_rent_prosp_verif_const, mortgage_retro_amt_const, mortgage_retro_verif_const, mortgage_prosp_amt_const, mortgage_prosp_verif_const, insurance_retro_amt_const, insurance_retro_verif_const, insurance_prosp_amt_const, insurance_prosp_verif_const, tax_retro_amt_const, tax_retro_verif_const, tax_prosp_amt_const, tax_prosp_verif_const, room_retro_amt_const, room_retro_verif_const, room_prosp_amt_const, room_prosp_verif_const, garage_retro_amt_const, garage_retro_verif_const, garage_prosp_amt_const, garage_prosp_verif_const, subsidy_retro_amt_const, subsidy_retro_verif_const, subsidy_prosp_amt_const, subsidy_prosp_verif_const, update_information_btn, save_information_btn, memb_btn_const, attempted_update_const, clear_all_btn, view_total_shel_btn)

		If page_to_display = HEST_dlg_page Then Call navigate_HEST_buttons(update_hest, err_msg, hest_update_attempted, update_information_btn, save_information_btn, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

		If ButtonPressed = ADDR_page_btn Then page_to_display = ADDR_dlg_page
		If ButtonPressed = SHEL_page_btn Then page_to_display = SHEL_dlg_page
		If ButtonPressed = HEST_page_btn Then page_to_display = HEST_dlg_page
	Loop until ButtonPressed = -1

	' If addr_update_attempted = True Then
	addr_eff_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year
	Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_addr_panel_info, addr_update_attempted)
	For shel_member = 0 to UBound(ALL_SHEL_PANELS_ARRAY, 2)
		' If ALL_SHEL_PANELS_ARRAY(attempted_update_const, shel_member) = True Then
		Call access_SHEL_panel("WRITE", ALL_SHEL_PANELS_ARRAY(shel_ref_number_const, shel_member), ALL_SHEL_PANELS_ARRAY(hud_sub_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(shared_yn_const, shel_member), ALL_SHEL_PANELS_ARRAY(paid_to_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(lot_rent_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(mortgage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(insurance_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(tax_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(room_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(garage_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_retro_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_amt_const, shel_member), ALL_SHEL_PANELS_ARRAY(subsidy_prosp_verif_const, shel_member), ALL_SHEL_PANELS_ARRAY(original_panel_info_const, shel_member))
		' End If
	Next
	If hest_update_attempted = True Then Call access_HEST_panel("WRITE", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)



End If


If select_option = "Change" Then

	Call read_total_SHEL_on_case(ref_numbers_with_panel, paid_to, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, total_shel_original_information)
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received, original_addr_panel_info, addr_update_attempted)
	Call access_HEST_panel("READ", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)

	housing_questions_step = 1
	view_addr_update_dlg = False
	view_shel_update_dlg = False
	view_shel_details_dlg = False
	page_to_display = CHNG_dlg_page

	Do
		err_msg = ""

		BeginDialog Dialog1, 0, 0, 555, 385, "Housing Expense Detail"

		  ButtonGroup ButtonPressed

			If page_to_display = CHNG_dlg_page Then
				Text 503, 57, 60, 10, "CHANGE"
				Call display_HOUSING_CHANGE_information(housing_questions_step, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, all_rent_verif, total_current_lot_rent, all_lot_rent_verif, total_current_garage, all_mortgage_verif, total_current_insurance, all_insurance_verif, total_current_taxes, all_taxes_verif, total_current_room, all_room_verif, total_current_mortgage, all_garage_verif, total_current_subsidy, all_subsidy_verif, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, housing_change_continue_btn, housing_change_overview_btn, housing_change_addr_update_btn, housing_change_shel_update_btn, housing_change_shel_details_btn, housing_change_review_btn)
			End If

			If page_to_display <> CHNG_dlg_page Then PushButton 485, 55, 65, 13, "CHANGE", CHNG_page_btn

			OkButton 450, 365, 50, 15
			CancelButton 500, 365, 50, 15

		EndDialog


		Dialog Dialog1
		cancel_without_confirmation

		If page_to_display = CHNG_dlg_page Then Call navigate_HOUSING_CHANGE_buttons(err_msg, housing_questions_step, household_move_yn, household_move_everyone_yn, move_date, shel_change_yn, shel_verif_received_yn, shel_start_date, shel_shared_yn, shel_subsidized_yn, total_current_rent, total_current_taxes, total_current_lot_rent, total_current_room, total_current_mortgage, total_current_garage, total_current_insurance, total_current_subsidy, shel_change_type, hest_heat_ac_yn, hest_electric_yn, hest_ac_on_electric_yn, hest_heat_on_electric_yn, hest_phone_yn, update_addr_button, addr_or_shel_change_notes, update_shel_button, housing_change_continue_btn, view_addr_update_dlg, view_shel_update_dlg, view_shel_details_dlg, addr_update_needed, shel_update_needed, hest_update_needed)


		If ButtonPressed = CHNG_page_btn Then page_to_display = CHNG_dlg_page
	Loop until ButtonPressed = -1




End If


script_end_procedure("Done")
