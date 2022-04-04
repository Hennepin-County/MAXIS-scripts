'GATHERING STATS===========================================================================================
name_of_script = "UTILITIES - WAIVED ER INTERVIEW SCREENING.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 270
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================
'run_locally = TRUE
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

function access_SHEL_panel(access_type, hud_sub_yn, shared_yn, paid_to, rent_retro_amt, rent_retro_verif, rent_prosp_amt, rent_prosp_verif, lot_rent_retro_amt, lot_rent_retro_verif, lot_rent_prosp_amt, lot_rent_prosp_verif, mortgage_retro_amt, mortgage_retro_verif, mortgage_prosp_amt, mortgage_prosp_verif, insurance_retro_amt, insurance_retro_verif, insurance_prosp_amt, insurance_prosp_verif, tax_retro_amt, tax_retro_verif, tax_prosp_amt, tax_prosp_verif, room_retro_amt, room_retro_verif, room_prosp_amt, room_prosp_verif, garage_retro_amt, garage_retro_verif, garage_prosp_amt, garage_prosp_verif, subsidy_retro_amt, subsidy_retro_verif, subsidy_prosp_amt, subsidy_prosp_verif)
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
        If rent_retro_verif = "OT" Then rent_retro_verif = "OT - Other Document"
        If rent_retro_verif = "NC" Then rent_retro_verif = "NC - Chg Rept, Neg Impact"
        If rent_retro_verif = "PC" Then rent_retro_verif = "PC - Chg Rept, Pos Imact"
        If rent_retro_verif = "NO" Then rent_retro_verif = "NO - No Ver Prvd"
        If rent_retro_verif = "__" Then rent_retro_verif = ""
        rent_prosp_amt = replace(rent_prosp_amt, "_", "")
        rent_prosp_amt = trim(rent_prosp_amt)
        If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
        If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
        If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
        If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Document"
        If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg Rept, Neg Impact"
        If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg Rept, Pos Imact"
        If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Ver Prvd"
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
        If lot_rent_retro_verif = "OT" Then lot_rent_retro_verif = "OT - Other Document"
        If lot_rent_retro_verif = "NC" Then lot_rent_retro_verif = "NC - Chg Rept, Neg Impact"
        If lot_rent_retro_verif = "PC" Then lot_rent_retro_verif = "PC - Chg Rept, Pos Imact"
        If lot_rent_retro_verif = "NO" Then lot_rent_retro_verif = "NO - No Ver Prvd"
        If lot_rent_retro_verif = "__" Then lot_rent_retro_verif = ""
        lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
        lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
        If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
        If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
        If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
        If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Document"
        If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg Rept, Neg Impact"
        If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg Rept, Pos Imact"
        If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Ver Prvd"
        If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""

        EMReadScreen mortgage_retro_amt,    8, 13, 37
        EMReadScreen mortgage_retro_verif,  2, 13, 48
        EMReadScreen mortgage_prosp_amt,    8, 13, 56
        EMReadScreen mortgage_prosp_verif,  2, 13, 67

        mortgage_retro_amt = replace(mortgage_retro_amt, "_", "")
        mortgage_retro_amt = trim(mortgage_retro_amt)
        If mortgage_retro_verif = "MO" Then mortgage_retro_verif = "MO - Mortgage Pmt Book"
        If mortgage_retro_verif = "CD" Then mortgage_retro_verif = "CD - Ctrct fro Deed"
        If mortgage_retro_verif = "OT" Then mortgage_retro_verif = "OT - Other Document"
        If mortgage_retro_verif = "NC" Then mortgage_retro_verif = "NC - Chg Rept, Neg Impact"
        If mortgage_retro_verif = "PC" Then mortgage_retro_verif = "PC - Chg Rept, Pos Imact"
        If mortgage_retro_verif = "NO" Then mortgage_retro_verif = "NO - No Ver Prvd"
        If mortgage_retro_verif = "__" Then mortgage_retro_verif = ""
        mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
        mortgage_prosp_amt = trim(mortgage_prosp_amt)
        If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
        If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
        If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Document"
        If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg Rept, Neg Impact"
        If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg Rept, Pos Imact"
        If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Ver Prvd"
        If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""

        EMReadScreen insurance_retro_amt,   8, 14, 37
        EMReadScreen insurance_retro_verif, 2, 14, 48
        EMReadScreen insurance_prosp_amt,   8, 14, 56
        EMReadScreen insurance_prosp_verif, 2, 14, 67

        insurance_retro_amt = replace(insurance_retro_amt, "_", "")
        insurance_retro_amt = trim(insurance_retro_amt)
        If insurance_retro_verif = "BI" Then insurance_retro_verif = "BI - Billing Stmt"
        If insurance_retro_verif = "OT" Then insurance_retro_verif = "OT - Other Document"
        If insurance_retro_verif = "NC" Then insurance_retro_verif = "NC - Chg Rept, Neg Impact"
        If insurance_retro_verif = "PC" Then insurance_retro_verif = "PC - Chg Rept, Pos Imact"
        If insurance_retro_verif = "NO" Then insurance_retro_verif = "NO - No Ver Prvd"
        If insurance_retro_verif = "__" Then insurance_retro_verif = ""
        insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
        insurance_prosp_amt = trim(insurance_prosp_amt)
        If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
        If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Document"
        If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg Rept, Neg Impact"
        If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg Rept, Pos Imact"
        If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Ver Prvd"
        If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""

        EMReadScreen tax_retro_amt,         8, 15, 37
        EMReadScreen tax_retro_verif,       2, 15, 48
        EMReadScreen tax_prosp_amt,         8, 15, 56
        EMReadScreen tax_prosp_verif,       2, 15, 67

        tax_retro_amt = replace(tax_retro_amt, "_", "")
        tax_retro_amt = trim(tax_retro_amt)
        If tax_retro_verif = "TX" Then tax_retro_verif = "TX - Prop Tax Stmt"
        If tax_retro_verif = "OT" Then tax_retro_verif = "OT - Other Document"
        If tax_retro_verif = "NC" Then tax_retro_verif = "NC - Chg Rept, Neg Impact"
        If tax_retro_verif = "PC" Then tax_retro_verif = "PC - Chg Rept, Pos Imact"
        If tax_retro_verif = "NO" Then tax_retro_verif = "NO - No Ver Prvd"
        If tax_retro_verif = "__" Then tax_retro_verif = ""
        tax_prosp_amt = replace(tax_prosp_amt, "_", "")
        tax_prosp_amt = trim(tax_prosp_amt)
        If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
        If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Document"
        If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg Rept, Neg Impact"
        If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg Rept, Pos Imact"
        If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Ver Prvd"
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
        If room_retro_verif = "OT" Then room_retro_verif = "OT - Other Document"
        If room_retro_verif = "NC" Then room_retro_verif = "NC - Chg Rept, Neg Impact"
        If room_retro_verif = "PC" Then room_retro_verif = "PC - Chg Rept, Pos Imact"
        If room_retro_verif = "NO" Then room_retro_verif = "NO - No Ver Prvd"
        If room_retro_verif = "__" Then room_retro_verif = ""
        room_prosp_amt = replace(room_prosp_amt, "_", "")
        room_prosp_amt = trim(room_prosp_amt)
        If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
        If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
        If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
        If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Document"
        If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg Rept, Neg Impact"
        If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg Rept, Pos Imact"
        If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Ver Prvd"
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
        If garage_retro_verif = "OT" Then garage_retro_verif = "OT - Other Document"
        If garage_retro_verif = "NC" Then garage_retro_verif = "NC - Chg Rept, Neg Impact"
        If garage_retro_verif = "PC" Then garage_retro_verif = "PC - Chg Rept, Pos Imact"
        If garage_retro_verif = "NO" Then garage_retro_verif = "NO - No Ver Prvd"
        If garage_retro_verif = "__" Then garage_retro_verif = ""
        garage_prosp_amt = replace(garage_prosp_amt, "_", "")
        garage_prosp_amt = trim(garage_prosp_amt)
        If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
        If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
        If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
        If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Document"
        If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg Rept, Neg Impact"
        If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg Rept, Pos Imact"
        If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Ver Prvd"
        If garage_prosp_verif = "__" Then garage_prosp_verif = ""

        EMReadScreen subsidy_retro_amt,     8, 18, 37
        EMReadScreen subsidy_retro_verif,   2, 18, 48
        EMReadScreen subsidy_prosp_amt,     8, 18, 56
        EMReadScreen subsidy_prosp_verif,   2, 18, 67

        subsidy_retro_amt = replace(subsidy_retro_amt, "_", "")
        subsidy_retro_amt = trim(subsidy_retro_amt)
        If subsidy_retro_verif = "SF" Then subsidy_retro_verif = "SF - Shelter Form"
        If subsidy_retro_verif = "LE" Then subsidy_retro_verif = "LE - Lease"
        If subsidy_retro_verif = "OT" Then subsidy_retro_verif = "OT - Other Document"
        If subsidy_retro_verif = "NO" Then subsidy_retro_verif = "NO - No Ver Prvd"
        If subsidy_retro_verif = "__" Then subsidy_retro_verif = ""
        subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
        subsidy_prosp_amt = trim(subsidy_prosp_amt)
        If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
        If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
        If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Document"
        If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Ver Prvd"
        If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""
    End If
end function

'===========================================================================================================================


'DECLARATIONS ==============================================================================================================
const memb_ref_numb                 = 00
const memb_last_name                = 01
const memb_first_name               = 02
const memb_age                      = 03
const clt_grh_status                = 04
const clt_hc_status                 = 05
const clt_snap_status               = 06
const memb_id_verif                 = 07
const memb_dob                      = 08
const memb_dob_verif                = 09
const memb_gender                   = 10
const memb_rel_to_applct            = 11
const memi_marriage_status          = 12
const memi_spouse_ref               = 13
const memi_spouse_name              = 14
const memi_citizen                  = 15
const memi_citizen_verif            = 16
const memi_last_grade               = 17
const memi_in_MN_less_12_mo         = 18
const memi_resi_verif               = 19
const memi_MN_entry_date            = 20
const disa_exists					= 21
const disa_begin_date 				= 22
const disa_end_date 				= 23
const schl_exists 					= 24
const schl_status					= 25
const schl_type						= 26
const memb_notes                    = 27


'===========================================================================================================================

EMConnect ""
Call check_for_MAXIS(TRUE)
Call MAXIS_case_number_finder(MAXIS_case_number)

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

recert_month = MonthName(DatePart("m", DateAdd("m", 1, date)))
recert_year = DatePart("yyyy", DateAdd("m", 1, date))
recert_year = recert_year & ""
If MAXIS_case_number <> "" Then
	Call Back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "REVW")
	EMReadScreen caf_date, 8, 13, 37
	If caf_date = "__ __ __" Then
		caf_date = ""
	Else
		caf_date = replace(caf_date, " ", "/")
	End If
	EMReadScreen SNAP_ER_code, 1, 7, 60
	If SNAP_ER_code <> "_" Then
		EMReadScreen snap_revw_type, 2, 9, 66
		If snap_revw_type = "ER" Then snap_er_checkbox = checked
	End If

	EMReadScreen CASH_ER_code, 1, 7, 40
	If CASH_ER_code <> "_" Then
		EMReadScreen cash_revw_type, 2, 9, 66
		If cash_revw_type = "ER" Then
			Call navigate_to_MAXIS_screen("STAT", "PROG")
			cash_one_prog = ""
			cash_two_prog = ""

			EMReadScreen cash_one_prog_status, 4, 6, 74
			EMReadScreen cash_two_prog_status, 4, 7, 74
			EMReadScreen grh_prog_status, 4, 9, 74

			If cash_one_prog_status = "ACTV" Then EMReadScreen cash_one_prog, 2, 6, 67
			If cash_two_prog_status = "ACTV" Then EMReadScreen cash_two_prog, 2, 6, 67
			If grh_prog_status = "ACTV" Then grh_er_checkbox = checked

			If cash_one_prog = "MF" OR cash_two_prog = "MF" Then mfip_er_checkbox = checked
			If cash_one_prog = "MS" OR cash_two_prog = "MS" Then msa_er_checkbox = checked
			If cash_one_prog = "GA" OR cash_two_prog = "GA" Then ga_er_checkbox = checked
		End If
	End If
End IF

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 256, 225, "CASE Information"
  EditBox 120, 35, 50, 15, MAXIS_case_number
  DropListBox 120, 60, 60, 45, "Select One..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", recert_month
  EditBox 190, 60, 35, 15, recert_year
  EditBox 120, 85, 50, 15, caf_date
  CheckBox 120, 105, 130, 10, "Check Here if there is no CAF in ECF", no_caf_checkbox
  CheckBox 70, 135, 30, 10, "SNAP", snap_er_checkbox
  CheckBox 110, 135, 30, 10, "MFIP", mfip_er_checkbox
  CheckBox 150, 135, 30, 10, "GRH", grh_er_checkbox
  CheckBox 185, 135, 25, 10, "GA", ga_er_checkbox
  CheckBox 215, 135, 25, 10, "MSA", msa_er_checkbox
  EditBox 80, 150, 170, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 190, 50, 15
    CancelButton 200, 205, 50, 15
  Text 35, 40, 80, 10, "Enter the Case Number:"
  Text 5, 65, 115, 10, "Enter the month of Recertification:"
  Text 20, 90, 95, 10, "Enter the CAF Date from ECF:"
  Text 5, 120, 150, 10, "Select ALL programs that are currently at ER:"
  Text 10, 155, 65, 10, "Worker Signature:"
  Text 10, 10, 230, 20, "This script will do a quick initial assessment to determine if this case is potentially eligible to have the recertification interview waived."
  Text 10, 175, 180, 10, "ADDITIONAL CASE PROCESSING WILL BE NEEDED."
  Text 10, 190, 180, 25, "Even if interview is waived a FULL case process and CASE:NOTE will be required. This script does NOT take all ER actions neeeded."
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		caf_date = trim(caf_date)

		Call validate_MAXIS_case_number(err_msg, "*")
		If recert_month = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the month of the ER."
		If len(recert_year) <> 2 AND len(recert_year) <> 4 Then err_msg = err_msg & vbNewLine & "* Enter the year of the ER as either 2 digits or four digits."
		If snap_er_checkbox = unchecked AND mfip_er_checkbox = unchecked AND grh_er_checkbox = unchecked AND ga_er_checkbox = unchecked AND msa_er_checkbox = unchecked THen  err_msg = err_msg & vbNewLine & "* Select all of the programs that have an ER for this months."
		If IsDate(caf_date) = FALSE and no_caf_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Enter the date of the CAF in ECF or indicate that there is no CAF in ECF by checking the box."
		If IsDate(caf_date) = TRUE and no_caf_checkbox = checked Then err_msg = err_msg & vbNewLine & "* You have indicated that there is a CAF by entering the CAF date but also checked the box indicating there is no CAF in ECF. Please select only one of these options."
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Sign your case note."

		If err_msg <> "" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

	Loop Until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

If recert_month = "January" Then MAXIS_footer_month = "01"
If recert_month = "February" Then MAXIS_footer_month = "02"
If recert_month = "March" Then MAXIS_footer_month = "03"
If recert_month = "April" Then MAXIS_footer_month = "04"
If recert_month = "May" Then MAXIS_footer_month = "05"
If recert_month = "June" Then MAXIS_footer_month = "06"
If recert_month = "July" Then MAXIS_footer_month = "07"
If recert_month = "August" Then MAXIS_footer_month = "08"
If recert_month = "September" Then MAXIS_footer_month = "09"
If recert_month = "October" Then MAXIS_footer_month = "10"
If recert_month = "November" Then MAXIS_footer_month = "11"
If recert_month = "December" Then MAXIS_footer_month = "12"
MAXIS_footer_year = right(recert_year, 2)

If grh_er_checkbox = checked OR ga_er_checkbox = checked OR msa_er_checkbox = checked Then
	If snap_er_checkbox = unchecked AND mfip_er_checkbox = unchecked Then
		end_msg = "*** This case - " & MAXIS_case_number & " - Only has an ER for an adult cash program. ***"
		end_msg = end_msg & vbCr & vbCr & "You did not indicate there was a SNAP or MFIP ER on this case, only an ER for:"
		If grh_er_checkbox = checked Then end_msg = end_msg & vbCr & "   - GRH"
		If ga_er_checkbox = checked Then end_msg = end_msg & vbCr & "   - GA"
		If msa_er_checkbox = checked Then end_msg = end_msg & vbCr & "   - MSA"
		end_msg = end_msg & vbCr & vbCr & "Adult Cash only reviews do not have an interview requirement and so does not need to be assessed to have it waived."
		end_msg = end_msg & vbCr & vbCr & "*** IF THERE ARE OTHER PROGRAMS AT ER FOR THIS CASE ***"
		end_msg = end_msg & vbCr & vbCr & "Rerun the sscript with the correct information entered into the dialog and the script will assess again."
		script_end_procedure_with_error_report(end_msg)
	End If
End If

run_another_script("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\reviews-delayed.vbs")
If recert_month = "October" Then
	REVW_ADJUSTED_ARRAY = oct_revw_to_adjust_array
	month_originally_waived = "APRIL"
End If
If recert_month = "November" Then
	REVW_ADJUSTED_ARRAY = nov_revw_to_adjust_array
	month_originally_waived = "MAY"
End If
If recert_month = "December" Then
	REVW_ADJUSTED_ARRAY = dec_revw_to_adjust_array
	month_originally_waived = "JUNE"
End If
If recert_month = "January" Then
	REVW_ADJUSTED_ARRAY = jan_revw_to_adjust_array
	month_originally_waived = "JULY"
End If
If recert_month = "February" Then
	REVW_ADJUSTED_ARRAY = feb_revw_to_adjust_array
	month_originally_waived = "AUGUST"
End If

For each revw_case in REVW_ADJUSTED_ARRAY
	If MAXIS_case_number = revw_case Then
		end_msg = "*** This case - " & MAXIS_case_number & " - has had the ER date ADJUSTED ***"
		end_msg = end_msg & vbCr & vbCr & "Any case that had their ER date adjusted in the months April - August due to the COVID Peacetime Emergency is not eligible to have the interview waived, they must complete an interview as a part of the recertification."
		end_msg = end_msg & vbCr & vbCr & "The script has checked the list from DHS of all the cases in MN that had the ER date waived in " & month_originally_waived & " and this case is on the list."
		end_msg = end_msg & vbCr & vbCr & "--- Process this case normally ---"
		script_end_procedure_with_error_report(end_msg)
	End If
Next
If no_caf_checkbox = checked Then
	end_msg = "*** There is no CAF on file ***"
	end_msg = end_msg & vbCr & vbCr & "In order to waive an interview for an ER we need to assess if the client has reported no changes on the CAF (or similar form) without this form, we cannot make an assessment."
	end_msg = end_msg & vbCr & vbCr & "--- If the work assigned is a callback for a 'phone CAF' or other callback, you must call the client and this will constitute an interview. Once we are talking to the client, this is considerded an interview and the full interview process must be followed.---"
	end_msg = end_msg & vbCr & vbCr & "If you misentered information, rerun the script and provide the CAF date."
	script_end_procedure_with_error_report(end_msg)
End If


BeginDialog Dialog1, 0, 0, 281, 175, "Client Request Call"
  DropListBox 80, 130, 190, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"Yes, called 3 times - no answer"+chr(9)+"No", call_answer
  ButtonGroup ButtonPressed
    OkButton 225, 155, 50, 15
  Text 75, 10, 120, 10, "--- We must return calls to clients ---"
  Text 10, 30, 250, 25, "If a client has called in, we must return the call. Once we have gotten a hold of the client, we are in the process of completing an interview and should process the case as normal. "
  Text 10, 65, 250, 10, "Do not use a possible waived interview as a reason to fail to return the call."
  Text 10, 85, 260, 20, "If you have made all efforts to contact the client for the callback and the client cannot be reached, you can review for processing without the interview."
  Text 5, 115, 255, 10, "Has the client or AREP called for an interview or requested a callback?"
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_confirmation

		If call_answer = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the client has called and is requesting a callback."

		If err_msg <> "" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

	Loop Until err_msg = ""
	If call_answer = "Yes" Then
		end_msg = "*** Clients who have requested an interview/callback need to be called ***"
		end_msg = end_msg & vbCr & vbCr & "Once we are on the phone with the client we are then completing an interview since we have reached the client and have the CAF."
		end_msg = end_msg & vbCr & vbCr & "Our goal is to maintain excellent customer service and that may include completing an interview."
		script_end_procedure_with_error_report(end_msg)
	End If
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'DIALOG WITH
	'which FORM was received
	'Address in MAXIS (homeless status)
	'HH comp

Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, living_situation_status, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, curr_phone_one, curr_phone_two, curr_phone_three, curr_phone_type_one, curr_phone_type_two, curr_phone_type_three)

Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "MEMB")

' Dim ALL_CLIENTS_ARRAY()
' ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)

member_counter = 0
Do
    EMReadScreen clt_ref_nbr, 2, 4, 33
    EMReadScreen clt_last_name, 25, 6, 30
    EMReadScreen clt_first_name, 12, 6, 63
    EMReadScreen clt_age, 3, 8, 76
	EMReadScreen id_verif, 2, 9, 68
	EMReadScreen date_of_birth, 10, 8, 42
	EMReadScreen dob_verif, 2, 8, 68
	EMReadScreen gender, 1, 9, 42
	EMReadScreen relastionship, 2, 10, 42

    ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, member_counter)
    ALL_CLIENTS_ARRAY(memb_ref_numb, member_counter) = clt_ref_nbr
    ALL_CLIENTS_ARRAY(memb_last_name, member_counter) = replace(clt_last_name, "_", "")
    ALL_CLIENTS_ARRAY(memb_first_name, member_counter) = replace(clt_first_name, "_", "")
    ALL_CLIENTS_ARRAY(memb_age, member_counter) = trim(clt_age)

	If id_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Birth Certificate"
	If id_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Religious Record"
	If id_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Drivers Lic/St ID"
	If id_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Divorce Decree"
	If id_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Alien Card"
	If id_verif = "AD" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Arrival/Depart"
	If id_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "\Doctor Stmt"
	If id_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "PPassport/Visa"
	If id_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "Other"
	If id_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_id_verif, member_counter) = "No Ver Prvd"
	ALL_CLIENTS_ARRAY(memb_dob, member_counter) = replace(date_of_birth, " ", "/")
	If dob_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Birth Certificate"
	If dob_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Religious Record"
	If dob_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Drivers Lic/St ID"
	If dob_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Divorce Decree"
	If dob_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Alien Card"
	If dob_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Doctor Stmt"
	If dob_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Passport/Visa"
	If dob_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "Other"
	If dob_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_dob_verif, member_counter) = "No Ver Prvd"
	If gender = "F" Then ALL_CLIENTS_ARRAY(memb_gender, member_counter) = "Female"
	If gender = "M" Then ALL_CLIENTS_ARRAY(memb_gender, member_counter) = "Male"
	If relastionship = "01" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Applicant"
    If relastionship = "02" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Spouse"
    If relastionship = "03" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Child"
    If relastionship = "04" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Parent"
    If relastionship = "05" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Sibling"
    If relastionship = "06" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Step Sibling"
    If relastionship = "08" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Step Child"
    If relastionship = "09" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Step Parent"
    If relastionship = "10" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Aunt"
    If relastionship = "11" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Uncle"
    If relastionship = "12" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Niece"
    If relastionship = "13" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Nephew"
    If relastionship = "14" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Cousin"
    If relastionship = "15" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Grandparent"
    If relastionship = "16" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Grandchild"
    If relastionship = "17" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Other Relative"
    If relastionship = "18" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Legal Guardian"
    If relastionship = "24" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Not Related"
    If relastionship = "25" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Live-In Attendant"
    If relastionship = "27" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, member_counter) = "Unknown"

    member_counter = member_counter + 1
    transmit
    EMReadScreen last_memb, 7, 24, 2
Loop until last_memb = "ENTER A"

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

    If clt_mar_status = "N" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Never married"
    If clt_mar_status = "M" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Married, Living with Spouse"
    If clt_mar_status = "S" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Married Living Apart"
    If clt_mar_status = "L" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Legally Separated"
    If clt_mar_status = "D" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Divorced"
    If clt_mar_status = "W" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Widowed"
    ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) = replace(clt_spouse, "_", "")
    If ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) <> "" Then
        For all_the_people = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
            If ALL_CLIENTS_ARRAY(memb_ref_nbr, all_the_people) = ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) Then
                ALL_CLIENTS_ARRAY(memi_spouse_name, case_memb) = ALL_CLIENTS_ARRAY(memb_first_name, all_the_people) & " " & ALL_CLIENTS_ARRAY(memb_last_name, all_the_people)
            End If
        Next
    End If
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
Next

Call navigate_to_MAXIS_screen("STAT", "DISA")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit

	EMReadScreen version_number, 1, 2, 73
	ALL_CLIENTS_ARRAY(disa_exists, case_memb) = FALSE
	If version_number = "1" Then
		ALL_CLIENTS_ARRAY(disa_exists, case_memb) = TRUE

		EMReadScreen begin_date, 10, 6, 47
		EMReadScreen end_date, 10, 6, 69

		ALL_CLIENTS_ARRAY(disa_begin_date, case_memb) = replace(begin_date, " ", "/")
		ALL_CLIENTS_ARRAY(disa_end_date, case_memb) = replace(end_date, " ", "/")
	End If
Next

Call navigate_to_MAXIS_screen("STAT", "SCHL")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit

	EMReadScreen version_number, 1, 2, 73
	ALL_CLIENTS_ARRAY(schl_exists, case_memb) = FALSE
	If version_number = "1" Then
		ALL_CLIENTS_ARRAY(schl_exists, case_memb) = TRUE

		EMReadScreen panel_status, 1, 6, 40
		EMReadScreen panel_type, 2, 7, 40

		If panel_status = "F" Then ALL_CLIENTS_ARRAY(schl_status, case_memb) = "Fulltime"
		If panel_status = "H" Then ALL_CLIENTS_ARRAY(schl_status, case_memb) = "At least halftime"
		If panel_status = "L" Then ALL_CLIENTS_ARRAY(schl_status, case_memb) = "Less than halftime"
		If panel_status = "N" Then ALL_CLIENTS_ARRAY(schl_status, case_memb) = "Not attending"

		If panel_type = "01" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "PreK - 6th"
		If panel_type = "11" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "7th  - 8th"
		If panel_type = "02" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "9th - 12th"
		If panel_type = "03" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "GED"
		If panel_type = "06" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "Child, not in school"
		If panel_type = "07" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "IEP"
		If panel_type = "08" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "Post-Secondary"
		If panel_type = "09" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "Grad Student"
		If panel_type = "10" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "Tech School"
		If panel_type = "12" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "Adult Basic Ed"
		If panel_type = "13" Then ALL_CLIENTS_ARRAY(schl_type, case_memb) = "ELL"
	End If
Next

dlg_len = 210 + UBound(ALL_CLIENTS_ARRAY, 2) * 15
grp_len = 65 + UBound(ALL_CLIENTS_ARRAY, 2) * 15

BeginDialog Dialog1, 0, 0, 541, dlg_len, "Household Information"
  Text 10, 15, 160, 10, "Which Recertification Form has been submitted?"
  DropListBox 175, 10, 150, 45, "Select One..."+chr(9)+"Combined Application Form (CAF)"+chr(9)+"MNbenefits"+chr(9)+"Combined Annual Renewal (CAR)", form_type_received
  Text 380, 5, 150, 10, "--- Open the Recertification Form from ECF ---"
  Text 380, 20, 150, 20, "The next questions require you to compare the form information with information from ECF."
  Text 380, 45, 150, 75, "The script will pull information from MAXIS to provide you an overview of information from the case as it is currently coded in MAXIS. While the script is running, you can still look into MAXIS directly to check for any information that the script does not pull into the dialogs. The script will also have a few questions about the information and the form to identify if there are changes reported in the ER form."

  GroupBox 10, 35, 360, 80, "MAXIS ADDRESS INFORMATION"
  Text 20, 50, 85, 10, "Homeless: " & addr_homeless
  Text 100, 50, 75, 10, "Residence Address:"
  Text 100, 65, 115, 10, resi_line_one
  If resi_line_two <> "" Then
	  Text 100, 75, 115, 10, resi_line_two
	  Text 100, 85, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
  Else
	  Text 100, 75, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
  End If
  Text 235, 50, 75, 10, "Mailing Address:"
  Text 235, 65, 115, 10, mail_line_one
  If mail_line_two <> "" Then
	  Text 235, 75, 115, 10, mail_line_two
	  Text 235, 85, 115, 10, mail_city & ", " & mail_state & " " & mail_zip
  Else
	  Text 235, 75, 115, 10, mail_city & ", " & mail_state & " " & mail_zip
  End If
  Text 20, 100, 250, 10, "Is there an address change or difference reported in the recertification form?"
  DropListBox 260, 95, 100, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", address_change_reported
  GroupBox 10, 120, 525, grp_len, "Household Composition"
  Text 20, 135, 110, 10, "Member Information"
  Text 275, 135, 40, 10, "Disability"
  Text 390, 135, 30, 10, "School"
  y_pos = 150
  For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	  Text 20, y_pos, 300, 10, "MEMB " & ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & "  -  " & ALL_CLIENTS_ARRAY(memb_first_name, case_memb) & ALL_CLIENTS_ARRAY(memb_last_name, case_memb) & "  - Age: " & ALL_CLIENTS_ARRAY(memb_age, case_memb) & "  -  Rel to applct: " & ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb)
	  If ALL_CLIENTS_ARRAY(disa_exists, case_memb) = TRUE Then
		  If IsDate(ALL_CLIENTS_ARRAY(disa_end_date, case_memb)) = TRUE Then
			  Text 275, y_pos, 120, 10, "Begin: " & ALL_CLIENTS_ARRAY(disa_begin_date, case_memb) & " - End:" & ALL_CLIENTS_ARRAY(disa_end_date, case_memb)
		  Else
			  Text 275, y_pos, 120, 10, "Begin: " & ALL_CLIENTS_ARRAY(disa_begin_date, case_memb) & " - No End Date"
		  End If
	  Else
		  Text 275, y_pos, 120, 10, "No DISA Panel"
	  End If

	  If ALL_CLIENTS_ARRAY(schl_exists, case_memb) = TRUE Then
		  Text 390, y_pos, 120, 10, "Status: " & ALL_CLIENTS_ARRAY(schl_status, case_memb) & " - Type: " &  ALL_CLIENTS_ARRAY(schl_type, case_memb)
	  Else
		  Text 390, y_pos, 120, 10, "No SCHL Panel"
	  End If
	  y_pos = y_pos + 15
  Next
  ' Text 20, 165, 300, 10, "MEMB 01  -  FIRST NAME  I. LAST NAME  - Age: xx  -  Relationship"
  ' Text 335, 165, 80, 10, "ACTIVE"
  ' Text 435, 165, 100, 10, "Attending - Elementary"
  ' Text 20, 180, 300, 10, "MEMB 01  -  FIRST NAME  I. LAST NAME  - Age: xx  -  Relationship"
  ' Text 335, 180, 80, 10, "Ended - mm/dd/yy"
  ' Text 435, 180, 100, 10, "Not Attending"
  Text 160, y_pos + 5, 265, 10, "Are there any differences or changes identified in the persons in the household?"
  DropListBox 425, y_pos, 100, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", hh_comp_change
  ButtonGroup ButtonPressed
    OkButton 430, dlg_len - 20, 50, 15
    CancelButton 485, dlg_len - 20, 50, 15
EndDialog

Do
	Do
		interview_needed = FALSE
		err_msg = ""

		dialog Dialog1
		cancel_confirmation

		If form_type_received = "Select One..." Then err_msg = err_msg & vbNewLine & "* List which form type has been received and is being reviewed."
		If address_change_reported = "Select One..." Then err_msg = err_msg & vbNewLine & "* Check the address on the form and the address listed for the case (the current panel information is listed in the dialog). Indicate if there are differences between the form and the current case reflecting a change."
		If hh_comp_change = "Select One..." Then err_msg = err_msg & vbNewLine & "* Check the members on the form and the household members listed for the case (the current panel information is listed in the dialog). Additionally check the disability and school status. Indicate if there are differences between the form and the current case reflecting a change."

		If err_msg <> "" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

	Loop Until err_msg = ""
	If address_change_reported = "CHANGES REPORTED" Then interview_needed = TRUE
	If hh_comp_change = "CHANGES REPORTED" Then interview_needed = TRUE

	If interview_needed = TRUE Then
		end_msg = "*** Changes / Differences Reported ***"
		end_msg = end_msg & vbCr & vbCr & "Since the information from the case appears to have potentially changed, we need to complete an interview."
		end_msg = end_msg & vbCr & vbCr & "Providing a full interview for a client with changes provides the best client service and ensures the most accuracy."
		script_end_procedure_with_error_report(end_msg)
	End If

	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE



const inc_category				= 00
const owner_name				= 01
const panel_call				= 02
const employer_const			= 03
const job_verif_const			= 04
const job_prosp_total_const		= 05
const job_prosp_hours_const		= 06
const job_freq_const			= 07
const fs_pic_freq_const			= 08
const fs_pic_ave_hours_const	= 09
const fs_pic_ave_pay_const		= 10
const fs_pic_monthly_inc_const	= 11
const busi_type_const			= 12
const se_method_const			= 13
const cash_net_prosp_amount		= 14
const snap_net_prosp_amount		= 15
const reported_hours			= 16
const unea_amt_const			= 17
const unea_type_const			= 18
const btn_const					= 19
' const
' const
' const
const income_notes				= 20


const shel_hud_sub_yn			= 00
const shel_shared_yn			= 01
const shel_paid_to				= 02
const shel_rent_retro_amt		= 03
const shel_rent_retro_verif		= 04
const shel_rent_prosp_amt		= 05
const shel_rent_prosp_verif		= 06
const shel_lot_rent_retro_amt	= 07
const shel_lot_rent_retro_verif	= 08
const shel_lot_rent_prosp_amt	= 09
const shel_lot_rent_prosp_verif	= 10
const shel_mortgage_retro_amt	= 11
const shel_mortgage_retro_verif	= 12
const shel_mortgage_prosp_amt	= 13
const shel_mortgage_prosp_verif	= 14
const shel_insurance_retro_amt	= 15
const shel_insurance_retro_verif= 16
const shel_insurance_prosp_amt	= 17
const shel_insurance_prosp_verif= 18
const shel_tax_retro_amt		= 19
const shel_tax_retro_verif		= 20
const shel_tax_prosp_amt		= 21
const shel_tax_prosp_verif		= 22
const shel_room_retro_amt		= 23
const shel_room_retro_verif		= 24
const shel_room_prosp_amt		= 25
const shel_room_prosp_verif		= 26
const shel_garage_retro_amt		= 27
const shel_garage_retro_verif	= 28
const shel_garage_prosp_amt		= 29
const shel_garage_prosp_verif	= 30
const shel_subsidy_retro_amt	= 31
const shel_subsidy_retro_verif	= 32
const shel_subsidy_prosp_amt	= 33
const shelter_member			= 34
const shel_subsidy_prosp_verif	= 35

const coex_ref_number 			= 0
Const support_verif_const		= 1
Const support_amount_const		= 2
Const alimony_verif_const		= 3
Const alimony_amount_const		= 4
Const tax_dep_verif_const		= 5
Const tax_dep_amount_const		= 6
Const other_verif_const			= 7
Const other_amount_const		= 8
const coex_notes				= 9

const dcex_ref_number		= 0
const dcex_instance			= 1
const provider_const 		= 2
const dcex_reason_const		= 3
const dcex_subsidy_const	= 4
const child_list_const		= 5
const total_amt_const		= 6
const dcex_verif_const		= 7
const dcex_notes 			= 8

Dim DCEX_ARRAY()
ReDim DCEX_ARRAY(dcex_notes, dcex_counter)


Dim INCOME_ARRAY()
ReDim INCOME_ARRAY(income_notes, 0)

Dim SHELTER_ARRAY()
ReDim SHELTER_ARRAY(shel_subsidy_prosp_verif, 0)

Dim COEX_ARRAY()
ReDim COEX_ARRAY(coex_notes, 0)

income_counter = 0
shel_count = 0
coex_counter = 0
dcex_counter = 0


For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
	Call navigate_to_MAXIS_screen("STAT", "JOBS")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve INCOME_ARRAY(income_notes, income_counter)

			INCOME_ARRAY(owner_name, income_counter) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " - " & ALL_CLIENTS_ARRAY(memb_first_name, case_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, case_memb)

            INCOME_ARRAY(inc_category, income_counter) = "JOBS"

			EMReadScreen version_number, 1, 2, 73
			INCOME_ARRAY(panel_call, income_counter) = "JOBS " & ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " 0" & version_number

			EMReadScreen employer_name, 30, 7, 42
			EMReadScreen job_verif, 1, 6, 34
			EMReadScreen job_prosp_total, 8, 17, 67
	        EMReadScreen job_frequency, 1, 18, 35
	        EMReadScreen job_prosp_hours, 3, 18, 72

			EMWriteScreen "X", 19, 38           'opening the FS PIC
			transmit

			EMReadScreen fs_pic_pay_frequency, 1, 5, 64
			EMReadScreen fs_pic_average_hours, 6, 16, 51
			EMReadScreen fs_pic_average_pay, 8, 17, 56
			EMReadScreen fs_pic_monthly_prospective, 8, 18, 56

			PF3                                 'closing the FS PIC


			INCOME_ARRAY(employer_const, income_counter) = replace(employer_name, "_", "")
			If job_verif = "1" Then INCOME_ARRAY(job_verif_const, income_counter) = "Pay Stubs"
	        If job_verif = "2" Then INCOME_ARRAY(job_verif_const, income_counter) = "Empl Stmt"
	        If job_verif = "3" Then INCOME_ARRAY(job_verif_const, income_counter) = "Coltrl Stmt"
	        If job_verif = "4" Then INCOME_ARRAY(job_verif_const, income_counter) = "Other Doc"
	        If job_verif = "5" Then INCOME_ARRAY(job_verif_const, income_counter) = "Pend Out State"
	        If job_verif = "N" Then INCOME_ARRAY(job_verif_const, income_counter) = "No Verif Prvd"
	        If job_verif = "?" Then INCOME_ARRAY(job_verif_const, income_counter) = "Delayed Verif"
			INCOME_ARRAY(job_prosp_total_const, income_counter) = trim(job_prosp_total)
	        INCOME_ARRAY(job_prosp_hours_const, income_counter) = trim(job_prosp_hours)
	        If job_frequency = "1" Then INCOME_ARRAY(job_freq_const, income_counter) = "1 - Monthly"
	        If job_frequency = "2" Then INCOME_ARRAY(job_freq_const, income_counter) = "2 - Semi Monthly"
	        If job_frequency = "3" Then INCOME_ARRAY(job_freq_const, income_counter) = "3 - Biweekly"
	        If job_frequency = "4" Then INCOME_ARRAY(job_freq_const, income_counter) = "4 -  Weekly"
	        If job_frequency = "5" Then INCOME_ARRAY(job_freq_const, income_counter) = "5 - Other"

			If fs_pic_pay_frequency = "1" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = "1 - Monthly"
	        If fs_pic_pay_frequency = "2" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = "2 - Semi Monthly"
	        If fs_pic_pay_frequency = "3" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = "3 - Biweekly"
	        If fs_pic_pay_frequency = "4" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = "4 -  Weekly"
	        If fs_pic_pay_frequency = "5" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = "5 - Other"
	        If fs_pic_pay_frequency = "_" Then INCOME_ARRAY(fs_pic_freq_const, income_counter) = ""

	        INCOME_ARRAY(fs_pic_ave_hours_const, income_counter) = trim(fs_pic_average_hours)
	        INCOME_ARRAY(fs_pic_ave_pay_const, income_counter) = trim(fs_pic_average_pay)
	        INCOME_ARRAY(fs_pic_monthly_inc_const, income_counter) = trim(fs_pic_monthly_prospective)

            income_counter = income_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

	Call navigate_to_MAXIS_screen("STAT", "BUSI")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit
	EMReadScreen versions, 6, 2, 73
	If versions <> "0 Of 0" Then
		Do
			ReDim Preserve INCOME_ARRAY(income_notes, income_counter)

			INCOME_ARRAY(owner_name, income_counter) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " - " & ALL_CLIENTS_ARRAY(memb_first_name, case_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, case_memb)

			INCOME_ARRAY(category_const, income_counter) = "BUSI"

			EMReadScreen version_number, 1, 2, 73
			INCOME_ARRAY(panel_call, income_counter) = "BUSI " & ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " 0" & version_number

			EMReadScreen type_of_income, 2, 5, 37
			EMReadScreen INCOME_ARRAY(cash_net_prosp_amount, income_counter), 8, 8, 69
			EMReadScreen INCOME_ARRAY(snap_net_prosp_amount, income_counter), 8, 10, 69
			EMReadScreen INCOME_ARRAY(reported_hours, income_counter), 3, 13, 74
			EMReadScreen SE_method, 2, 16, 53

			If type_of_income = "01" Then INCOME_ARRAY(busi_type_const, income_counter) = "Farming"
	        If type_of_income = "02" Then INCOME_ARRAY(busi_type_const, income_counter) = "Real Estate"
	        If type_of_income = "03" Then INCOME_ARRAY(busi_type_const, income_counter) = "Home Product Sales"
	        If type_of_income = "04" Then INCOME_ARRAY(busi_type_const, income_counter) = "Other Sales"
	        If type_of_income = "05" Then INCOME_ARRAY(busi_type_const, income_counter) = "Personal Services"
	        If type_of_income = "06" Then INCOME_ARRAY(busi_type_const, income_counter) = "Paper Route"
	        If type_of_income = "07" Then INCOME_ARRAY(busi_type_const, income_counter) = "In Home Daycare"
	        If type_of_income = "08" Then INCOME_ARRAY(busi_type_const, income_counter) = "Rental Income"
	        If type_of_income = "09" Then INCOME_ARRAY(busi_type_const, income_counter) = "Other"
			If SE_method = "01" Then INCOME_ARRAY(se_method_const, income_counter) = "50% Gross Inc"
	        If SE_method = "02" Then INCOME_ARRAY(se_method_const, income_counter) = "Tax Forms"
			INCOME_ARRAY(cash_net_prosp_amount, income_counter) = trim(INCOME_ARRAY(cash_net_prosp_amount, income_counter))
			INCOME_ARRAY(snap_net_prosp_amount, income_counter) = trim(INCOME_ARRAY(snap_net_prosp_amount, income_counter))

			income_counter = income_counter + 1
			transmit
			EMReadScreen last_panel, 7, 24, 2
		Loop until last_panel = "ENTER A"
	End If

	Call navigate_to_MAXIS_screen("STAT", "UNEA")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit
	EMReadScreen versions, 6, 2, 73
	If versions <> "0 Of 0" Then
		Do
			ReDim Preserve INCOME_ARRAY(income_notes, income_counter)

			INCOME_ARRAY(owner_name, income_counter) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " - " & ALL_CLIENTS_ARRAY(memb_first_name, case_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, case_memb)

			INCOME_ARRAY(category_const, income_counter) = "UNEA"

			EMReadScreen version_number, 1, 2, 73
			INCOME_ARRAY(panel_call, income_counter) = "UNEA " & ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb) & " 0" & version_number

			EMReadScreen panel_type, 2, 5, 37
			EMReadScreen total_amount, 8, 18, 68

			If panel_type = "01" Then INCOME_ARRAY(unea_type_const, income_counter) = "RSDI, Disa"
			If panel_type = "02" Then INCOME_ARRAY(unea_type_const, income_counter) = "RSDI, No Disa"
			If panel_type = "03" Then INCOME_ARRAY(unea_type_const, income_counter) = "SSI"
			If panel_type = "06" Then INCOME_ARRAY(unea_type_const, income_counter) = "Non-MN PA"
			If panel_type = "11" Then INCOME_ARRAY(unea_type_const, income_counter) = "VA Disability"
			If panel_type = "12" Then INCOME_ARRAY(unea_type_const, income_counter) = "VA Pension"
			If panel_type = "13" Then INCOME_ARRAY(unea_type_const, income_counter) = "VA Other"
			If panel_type = "38" Then INCOME_ARRAY(unea_type_const, income_counter) = "VA Aid & Attendance"
			If panel_type = "14" Then INCOME_ARRAY(unea_type_const, income_counter) = "Unemployment Insurance"
			If panel_type = "15" Then INCOME_ARRAY(unea_type_const, income_counter) = "Worker's Comp"
			If panel_type = "16" Then INCOME_ARRAY(unea_type_const, income_counter) = "Railroad Retirement"
			If panel_type = "17" Then INCOME_ARRAY(unea_type_const, income_counter) = "Other Retirement"
			If panel_type = "18" Then INCOME_ARRAY(unea_type_const, income_counter) = "Military Enrirlement"
			If panel_type = "19" Then INCOME_ARRAY(unea_type_const, income_counter) = "FC Child req FS"
			If panel_type = "20" Then INCOME_ARRAY(unea_type_const, income_counter) = "FC Child not req FS"
			If panel_type = "21" Then INCOME_ARRAY(unea_type_const, income_counter) = "FC Adult req FS"
			If panel_type = "22" Then INCOME_ARRAY(unea_type_const, income_counter) = "FC Adult not req FS"
			If panel_type = "23" Then INCOME_ARRAY(unea_type_const, income_counter) = "Dividends"
			If panel_type = "24" Then INCOME_ARRAY(unea_type_const, income_counter) = "Interest"
			If panel_type = "25" Then INCOME_ARRAY(unea_type_const, income_counter) = "Cnt gifts/prizes"
			If panel_type = "26" Then INCOME_ARRAY(unea_type_const, income_counter) = "Strike Benefits"
			If panel_type = "27" Then INCOME_ARRAY(unea_type_const, income_counter) = "Contract for Deed"
			If panel_type = "28" Then INCOME_ARRAY(unea_type_const, income_counter) = "Illegal Income"
			If panel_type = "29" Then INCOME_ARRAY(unea_type_const, income_counter) = "Other Countable"
			If panel_type = "30" Then INCOME_ARRAY(unea_type_const, income_counter) = "Infrequent"
			If panel_type = "31" Then INCOME_ARRAY(unea_type_const, income_counter) = "Other - FS Only"
			If panel_type = "08" Then INCOME_ARRAY(unea_type_const, income_counter) = "Direct Child Support"
			If panel_type = "35" Then INCOME_ARRAY(unea_type_const, income_counter) = "Direct Spousal Support"
			If panel_type = "36" Then INCOME_ARRAY(unea_type_const, income_counter) = "Disbursed Child Support"
			If panel_type = "37" Then INCOME_ARRAY(unea_type_const, income_counter) = "Disbursed Spousal Support"
			If panel_type = "39" Then INCOME_ARRAY(unea_type_const, income_counter) = "Disbursed CS Arrears"
			If panel_type = "40" Then INCOME_ARRAY(unea_type_const, income_counter) = "Disbursed Spsl Sup Arrears"
			If panel_type = "43" Then INCOME_ARRAY(unea_type_const, income_counter) = "Disbursed Excess CS"
			If panel_type = "44" Then INCOME_ARRAY(unea_type_const, income_counter) = "MSA - Excess Income for SSI"
			If panel_type = "47" Then INCOME_ARRAY(unea_type_const, income_counter) = "Tribal Income"
			If panel_type = "48" Then INCOME_ARRAY(unea_type_const, income_counter) = "Trust Income"
			If panel_type = "49" Then INCOME_ARRAY(unea_type_const, income_counter) = "Non-Recurring"
			INCOME_ARRAY(unea_amt_const, income_counter) = trim(total_amount)

			income_counter = income_counter + 1
			transmit
			EMReadScreen last_panel, 7, 24, 2
		Loop until last_panel = "ENTER A"
	End If

	Call navigate_to_MAXIS_screen("STAT", "SHEL")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit
	EMReadScreen versions, 6, 2, 73
	If versions <> "0 Of 0" Then
		ReDim Preserve SHELTER_ARRAY(shel_subsidy_prosp_verif, shel_count)
		call access_SHEL_panel("READ", SHELTER_ARRAY(shel_hud_sub_yn, shel_count), SHELTER_ARRAY(shel_shared_yn, shel_count), SHELTER_ARRAY(shel_paid_to, shel_count), SHELTER_ARRAY(shel_rent_retro_amt, shel_count), SHELTER_ARRAY(shel_rent_retro_verif, shel_count), SHELTER_ARRAY(shel_rent_prosp_amt, shel_count), SHELTER_ARRAY(shel_rent_prosp_verif, shel_count), SHELTER_ARRAY(shel_lot_rent_retro_amt, shel_count), SHELTER_ARRAY(shel_lot_rent_retro_verif, shel_count), SHELTER_ARRAY(shel_lot_rent_prosp_amt, shel_count), SHELTER_ARRAY(shel_lot_rent_prosp_verif, shel_count), SHELTER_ARRAY(shel_mortgage_retro_amt, shel_count), SHELTER_ARRAY(shel_mortgage_retro_verif, shel_count), SHELTER_ARRAY(shel_mortgage_prosp_amt, shel_count), SHELTER_ARRAY(shel_mortgage_prosp_verif, shel_count), SHELTER_ARRAY(shel_insurance_retro_amt, shel_count), SHELTER_ARRAY(shel_insurance_retro_verif, shel_count), SHELTER_ARRAY(shel_insurance_prosp_amt, shel_count), SHELTER_ARRAY(shel_insurance_prosp_verif, shel_count), SHELTER_ARRAY(shel_tax_retro_amt, shel_count), SHELTER_ARRAY(shel_tax_retro_verif, shel_count), SHELTER_ARRAY(shel_tax_prosp_amt, shel_count), SHELTER_ARRAY(shel_tax_prosp_verif, shel_count), SHELTER_ARRAY(shel_room_retro_amt, shel_count), SHELTER_ARRAY(shel_room_retro_verif, shel_count), SHELTER_ARRAY(shel_room_prosp_amt, shel_count), SHELTER_ARRAY(shel_room_prosp_verif, shel_count), SHELTER_ARRAY(shel_garage_retro_amt, shel_count), SHELTER_ARRAY(shel_garage_retro_verif, shel_count), SHELTER_ARRAY(shel_garage_prosp_amt, shel_count), SHELTER_ARRAY(shel_garage_prosp_verif, shel_count), SHELTER_ARRAY(shel_subsidy_retro_amt, shel_count), SHELTER_ARRAY(shel_subsidy_retro_verif, shel_count), SHELTER_ARRAY(shel_subsidy_prosp_amt, shel_count), SHELTER_ARRAY(shel_subsidy_prosp_verif, shel_count))

		SHELTER_ARRAY(shelter_member, shel_count) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb)
		shel_count = shel_count + 1
	End If

	Call navigate_to_MAXIS_screen("STAT", "COEX")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit
	EMReadScreen versions, 6, 2, 73
	If versions <> "0 Of 0" Then

		ReDim Preserve COEX_ARRAY(coex_notes, coex_counter)

		EMReadScreen support_verif, 1, 10, 36
		EMReadScreen support_amount, 8, 10, 63
		EMReadScreen alimony_verif, 1, 11, 36
		EMReadScreen alimony_amount, 8, 11, 63
		EMReadScreen tax_dep_verif, 1, 12, 36
		EMReadScreen tax_dep_amount, 8, 12, 63
		EMReadScreen other_verif, 1, 13, 36
		EMReadScreen other_amount, 8, 13, 63

		COEX_ARRAY(coex_ref_number, coex_counter) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb)
		If support_verif = "1" Then COEX_ARRAY(support_verif_const, coex_counter) = "Canceled Checks/Money Orders"
		If support_verif = "2" Then COEX_ARRAY(support_verif_const, coex_counter) = "Receipts"
		If support_verif = "3" Then COEX_ARRAY(support_verif_const, coex_counter) = "Colateral Statement"
		If support_verif = "4" Then COEX_ARRAY(support_verif_const, coex_counter) = "Other Document"
		If support_verif = "N" Then COEX_ARRAY(support_verif_const, coex_counter) = "No Verif"
		If support_verif = "_" Then COEX_ARRAY(support_verif_const, coex_counter) = ""
		COEX_ARRAY(support_amount_const, coex_counter) = trim(replace(support_amount, "_", ""))
		' COEX_ARRAY(alimony_verif_const, coex_counter) =
		If support_verif = "1" Then COEX_ARRAY(alimony_verif_const, coex_counter) = "Canceled Checks/Money Orders"
		If support_verif = "2" Then COEX_ARRAY(alimony_verif_const, coex_counter) = "Receipts"
		If support_verif = "3" Then COEX_ARRAY(alimony_verif_const, coex_counter) = "Colateral Statement"
		If support_verif = "4" Then COEX_ARRAY(alimony_verif_const, coex_counter) = "Other Document"
		If support_verif = "N" Then COEX_ARRAY(alimony_verif_const, coex_counter) = "No Verif"
		If support_verif = "_" Then COEX_ARRAY(alimony_verif_const, coex_counter) = ""
		COEX_ARRAY(alimony_amount_const, coex_counter) = trim(replace(alimony_amount, "_", ""))
		' COEX_ARRAY(tax_dep_verif_const, coex_counter) =
		If support_verif = "1" Then COEX_ARRAY(tax_dep_verif_const, coex_counter) = "Tax Form"
		If support_verif = "2" Then COEX_ARRAY(tax_dep_verif_const, coex_counter) = "Colateral Statement"
		If support_verif = "N" Then COEX_ARRAY(tax_dep_verif_const, coex_counter) = "No Verif"
		If support_verif = "_" Then COEX_ARRAY(tax_dep_verif_const, coex_counter) = ""
		COEX_ARRAY(tax_dep_amount_const, coex_counter) = trim(replace(tax_dep_amount, "_", ""))
		' COEX_ARRAY(other_verif_const, coex_counter) =
		If support_verif = "1" Then COEX_ARRAY(other_verif_const, coex_counter) = "Canceled Checks/Money Orders"
		If support_verif = "2" Then COEX_ARRAY(other_verif_const, coex_counter) = "Receipts"
		If support_verif = "3" Then COEX_ARRAY(other_verif_const, coex_counter) = "Colateral Statement"
		If support_verif = "4" Then COEX_ARRAY(other_verif_const, coex_counter) = "Other Document"
		If support_verif = "N" Then COEX_ARRAY(other_verif_const, coex_counter) = "No Verif"
		If support_verif = "_" Then COEX_ARRAY(other_verif_const, coex_counter) = ""
		COEX_ARRAY(other_amount_const, coex_counter) = trim(replace(other_amount, "_", ""))


		coex_counter = coex_counter + 1
	End If

	Call navigate_to_MAXIS_screen("STAT", "DCEX")
	EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
	transmit
	EMReadScreen versions, 6, 2, 73
	If versions <> "0 Of 0" Then
		Do
			ReDim Preserve DCEX_ARRAY(dcex_notes, dcex_counter)

			EMReadScreen version_number, 1, 2, 73

			DCEX_ARRAY(dcex_ref_number, dcex_counter) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb)
			DCEX_ARRAY(dcex_instance, dcex_counter) = "0" & version_number

			EMReadScreen provider, 25, 6, 47
			EMReadScreen reason_code, 1, 7, 44
			EMReadScreen subsidy_code, 1, 8, 44

			DCEX_ARRAY(provider_const, dcex_counter) = trim(replace(provider, "_", ""))
			If reason_code = "J" Then DCEX_ARRAY(dcex_reason_const, dcex_counter) = "Job"
			If reason_code = "S" Then DCEX_ARRAY(dcex_reason_const, dcex_counter) = "School/Training"
			If reason_code = "L" Then DCEX_ARRAY(dcex_reason_const, dcex_counter) = "Looking for Work"
			If reason_code = "O" Then DCEX_ARRAY(dcex_reason_const, dcex_counter) = "Other"

			If subsidy_code = "B" Then DCEX_ARRAY(dcex_subsidy_const, dcex_counter) = "Basic Sliding Fee"
			If subsidy_code = "M" Then DCEX_ARRAY(dcex_subsidy_const, dcex_counter) = "MFIP Child Care"
			If subsidy_code = "P" Then DCEX_ARRAY(dcex_subsidy_const, dcex_counter) = "Post Secondary"
			If subsidy_code = "T" Then DCEX_ARRAY(dcex_subsidy_const, dcex_counter) = "Transition Year Child Care"
			If subsidy_code = "_" Then DCEX_ARRAY(dcex_subsidy_const, dcex_counter) = "None"

			payment_subtotal = 0
			dcex_row = 11
			Do
				EMReadScreen child_ref, 2, dcex_row, 29
				EMReadScreen verif_code, 1, dcex_row, 41
				EMReadScreen prosp_amt, 8, dcex_row, 63

				If child_ref <> "__" Then
					DCEX_ARRAY(child_list_const, dcex_counter) = DCEX_ARRAY(child_list_const, dcex_counter)  & child_ref & ", "
					prosp_amt = trim(replace(prosp_amt, "_", ""))
					If IsNumeric(prosp_amt) = TRUE Then payment_subtotal = payment_subtotal + prosp_amt
				End If

				dcex_row = dcex_row + 1
				If dcex_row = 17 Then
					PF20
					EMReadScreen end_of_list, 9, 24, 14
					If end_of_list = "LAST PAGE" Then Exit Do
					dcex_row = 11
				End If
			Loop Until child_ref = "__"

			If right(DCEX_ARRAY(child_list_const, dcex_counter), 2) = ", " Then DCEX_ARRAY(child_list_const, dcex_counter) = left(DCEX_ARRAY(child_list_const, dcex_counter), len(DCEX_ARRAY(child_list_const, dcex_counter)) - 2)
			DCEX_ARRAY(total_amt_const, dcex_counter) = payment_subtotal & ""
			If verif_code = "1" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = "Child Care Verif Form"
			If verif_code = "2" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = "Cancelled Checks/Receipts"
			If verif_code = "3" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = "Provider Statement"
			If verif_code = "4" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = "Other"
			If verif_code = "N" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = "No Verif"
			If verif_code = "_" Then DCEX_ARRAY(dcex_verif_const, dcex_counter) = ""

			dcex_counter = dcex_counter + 1
			transmit
			EMReadScreen last_panel, 7, 24, 2
		Loop until last_panel = "ENTER A"
	End If
Next


Call navigate_to_MAXIS_screen("STAT", "HEST")

EMReadScreen Heat_Air_YN, 1, 13, 60
EMReadScreen Heat_Air_Amount, 6, 13, 75
EMReadScreen Electric_YN, 1, 14, 60
EMReadScreen Electric_Amount, 6, 14, 75
EMReadScreen Phone_YN, 1, 15, 60
EMReadScreen Phone_Amount, 6, 15, 75

If Heat_Air_YN = "Y" Then Heat_Air_YN = "Yes"
If Heat_Air_YN = "N" or Heat_Air_YN = "_" Then Heat_Air_YN = "No"
Heat_Air_Amount = trim(Heat_Air_Amount)
If Electric_YN = "Y" Then Electric_YN = "Yes"
If Electric_YN = "N" or Electric_YN = "_" Then Electric_YN = "No"
Electric_Amount = trim(Electric_Amount)
If Phone_YN = "Y" Then Phone_YN = "Yes"
If Phone_YN = "N" or Phone_YN = "_" Then Phone_YN = "No"
Phone_Amount = trim(Phone_Amount)

' EMReadScreen


'IF Combined AR
	'Review information on Q 2
	'List Income & Expenses - ask if changes are apparent
If form_type_received = "Combined Annual Renewal (CAR)" Then
	dlg_len = 200
	earned_grp = 20
	unearned_grp = 20
	housing_grp_len = 35
	other_grp_len = 40

	UNEA_exists = TRUE
	EARNED_exists = TRUE
	shel_exists = FALSE
	dcex_exists = FALSE
	coex_exists = FALSE

	y_pos = 45

	For each_income = 0 to UBound(INCOME_ARRAY, 2)
		If INCOME_ARRAY(category_const, each_income) = "JOBS" Then earned_grp = earned_grp + 10
		If INCOME_ARRAY(category_const, each_income) = "BUSI" Then earned_grp = earned_grp + 10
		If INCOME_ARRAY(category_const, each_income) = "UNEA" Then unearned_grp = unearned_grp + 10

		dlg_len = dlg_len + 10
	Next
	If earned_grp = 20 Then
		earned_grp = 30
		EARNED_exists = FALSE
	End If
	If unearned_grp = 20 Then
		unearned_grp = 30
		UNEA_exists = FALSE
	End If

	For case_shel = 0 To UBound(SHELTER_ARRAY, 2)
		If SHELTER_ARRAY(shelter_member, case_shel) <> "" Then
			shel_exists = TRUE
			If SHELTER_ARRAY(shel_rent_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_tax_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_room_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_garage_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
		End If
	Next

	For case_dcex = 0 to UBound(DCEX_ARRAY, 2)
		If DCEX_ARRAY(dcex_ref_number, case_dcex) <> "" Then
			dcex_exists = TRUE
			other_grp_len = other_grp_len + 10
			dlg_len = dlg_len + 10
		End If
	Next

	For case_coex = 0 to UBound(COEX_ARRAY, 2)
		If COEX_ARRAY(coex_ref_number, case_coex) <> "" Then
			coex_exists = TRUE
			other_grp_len = other_grp_len + 10
			dlg_len = dlg_len + 10
		End If
	Next

	' MsgBox earned_grp
	BeginDialog Dialog1, 0, 0, 616, dlg_len, "Case Review"
	  ButtonGroup ButtonPressed
	  Text 10, 15, 285, 10, "Check question 2 on the Combined AR Form. Does this question indicate any changes?"
	  DropListBox 300, 10, 120, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", quest_two_changes_reported
	  GroupBox 10, 30, 600, earned_grp, "Earned Income in MAXIS"
	  If EARNED_exists = FALSE Then
		  Text 20, y_pos, 200, 10, "There is no EARNED INCOME listed on this case in MAXIS."
		  y_pos = y_pos + 10
	  Else
		  For each_income = 0 to UBound(INCOME_ARRAY, 2)
			  If INCOME_ARRAY(category_const, each_income) = "JOBS" Then
				  PushButton 20, y_pos, 45, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 115, 10, "Employer: " & INCOME_ARRAY(employer_const, each_income)
				  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(job_prosp_total_const, each_income)
				  Text 335, y_pos, 45, 10, "Hours:" & INCOME_ARRAY(job_prosp_hours_const, each_income)
				  Text 395, y_pos, 120, 10, "Pay Frequency: " & INCOME_ARRAY(job_freq_const, each_income)
				  y_pos = y_pos + 10
			  End If

			  If INCOME_ARRAY(category_const, each_income) = "BUSI" Then
				  PushButton 20, y_pos, 45, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 115, 10, "Self Employment Type: " & INCOME_ARRAY(busi_type_const, each_income)
				  If INCOME_ARRAY(cash_net_prosp_amount, each_income) <> "0.00" Then
					  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(cash_net_prosp_amount, each_income)
				  Else
					  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(snap_net_prosp_amount, each_income)
				  End If

				  Text 335, y_pos, 45, 10, "Hours:" & INCOME_ARRAY(reported_hours, each_income)
				  y_pos = y_pos + 10
			  End If
		  Next
	  End If
	  y_pos = y_pos + 10
	  GroupBox 10, y_pos, 600, unearned_grp, "Unearned Income in MAXIS"
	  y_pos = y_pos +15
	  If UNEA_exists = FALSE Then
		  Text 20, y_pos, 200, 10, "There is no UNEARNED INCOME listed on this case in MAXIS."
		  y_pos = y_pos + 10
	  Else
		  For each_income = 0 to UBound(INCOME_ARRAY, 2)
			  If INCOME_ARRAY(category_const, each_income) = "UNEA" Then
				  PushButton 20, y_pos, 50, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 110, 10, "Source: " & INCOME_ARRAY(unea_type_const, each_income)
				  Text 210, y_pos, 105, 10, "Income Amount: $" & INCOME_ARRAY(unea_amt_const, each_income)
				  y_pos = y_pos + 10
			  End If

		  Next
	  End If
	   y_pos = y_pos + 10
	   GroupBox 10, y_pos, 600, housing_grp_len, "Housing Expenses"
	   y_pos = y_pos + 15
 	  Text 20, y_pos, 50, 10, "Shelter:"
 	  If shel_exists = FALSE Then
 		  Text 25, y_pos, 200, 10, "There is no Shelter Expenses Listed in MAXIS"
 		  y_pos = y_pos + 10
 	  Else
 		  For case_shel = 0 To UBound(SHELTER_ARRAY, 2)
 		  	  If SHELTER_ARRAY(shelter_member, case_shel) <> "" Then
 				  Text 75, y_pos, 125, 10, "Landlord: " & SHELTER_ARRAY(shel_paid_to, case_shel)
 				  Text 445, y_pos, 100, 10, "Listed under: MEMB " & SHELTER_ARRAY(shelter_member, case_shel)
 					If SHELTER_ARRAY(shel_rent_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Rent: $" & SHELTER_ARRAY(shel_rent_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_rent_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Lot Rent: $" & SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_lot_rent_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Mortgage: $" & SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_mortgage_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Insurance: $" & SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_insurance_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_tax_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Tax: $" & SHELTER_ARRAY(shel_tax_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_tax_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_room_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Room: $" & SHELTER_ARRAY(shel_room_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_room_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_garage_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Garage: $" & SHELTER_ARRAY(shel_garage_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_garage_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 					If SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel) <> "" Then
 						Text 205, y_pos, 70, 10, "Subsidy: $" & SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel)
 						Text 290, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_subsidy_prosp_verif, case_shel)
 						y_pos = y_pos + 10
 					End If
 			  End If
 		  Next
 	  End If
 	  y_pos = y_pos + 5
 	  Text 20, y_pos, 35, 10, "Utilities:"
 	  Text 75, y_pos, 60, 10, "Heat/AC: " & Heat_Air_YN
 	  Text 135, y_pos, 70, 10, "Amount - $" & Heat_Air_Amount
 	  Text 250, y_pos, 60, 10, "Electric: " & Electric_YN
 	  Text 310, y_pos, 70, 10, "Amount - $" & Electric_Amount
 	  Text 415, y_pos, 60, 10, "Phone: " & Phone_YN
 	  Text 475, y_pos, 70, 10, "Amount - $" & Phone_Amount
 	  y_pos = y_pos + 20
 	  GroupBox 10, y_pos, 600, other_grp_len, "Other Expenses"
 	  y_pos = y_pos + 10
 	  Text 20, y_pos, 90, 10, "Dependent Care Expense:"
 	  If dcex_exists = FALSE Then
 		  Text 120, y_pos, 200, 10, "There are no DEPENDENT CARE EXPENSES listed in MAXIS."
 		  y_pos = y_pos + 10
 	  Else
	  	y_pos = y_pos + 10
 		  For case_dcex = 0 to UBound(DCEX_ARRAY, 2)
 	  		  If DCEX_ARRAY(dcex_ref_number, case_dcex) <> "" Then
 				  Text 25, y_pos, 110, 10, "Provider: " & DCEX_ARRAY(provider_const, case_dcex)
 				  Text 140, y_pos, 105, 10, "For: " & DCEX_ARRAY(child_list_const, case_dcex)
 				  Text 255, y_pos, 65, 10, "Amount: $" & DCEX_ARRAY(total_amt_const, case_dcex)
 				  Text 330, y_pos, 90, 10, "Verif: " & DCEX_ARRAY(dcex_verif_const, case_dcex)
 				  Text 425, y_pos, 95, 10, "Reason: " & DCEX_ARRAY(dcex_reason_const, case_dcex)
 				  y_pos = y_pos + 10
 			  End If
 		  Next
 	  End If
 	  y_pos = y_pos + 5
 	  Text 20, y_pos, 90, 10, "Court Ordered Expense:"
 	  If coex_exists = FALSE Then
 		  Text 120, y_pos, 200, 10, "There are no COURT ORDERED EXPENSES listed in MAXIS."
 		  y_pos = y_pos + 10
 	  Else
		  y_pos = y_pos + 10
 		  For case_coex = 0 to UBound(COEX_ARRAY, 2)
 		  	  If COEX_ARRAY(coex_ref_number, case_coex) <> "" Then
 				  Text 25, y_pos, 70, 10, "Paid by: MEMB " & COEX_ARRAY(coex_ref_number, case_coex)
 				  Text 100, y_pos, 80, 10, "Support: Amt - $" & COEX_ARRAY(support_amount_const, case_coex)
 				  Text 185, y_pos, 80, 10, "Alimony: Amt - $" & COEX_ARRAY(alimony_amount_const, case_coex)
 				  Text 270, y_pos, 80, 10, "Tax Dep: Amt - $" & COEX_ARRAY(tax_dep_amount_const, case_coex)
 				  Text 355, y_pos, 80, 10, "Other: Amt - $" & COEX_ARRAY(other_amount_const, case_coex)
 				  y_pos = y_pos + 10
 			  End If
 		  Next
 	  End If
 	  y_pos = y_pos + 15
	  Text 15, y_pos, 285, 10, "Is there any information in the CAR Form that indicates a change to any of this detail?"
	  DropListBox 300, y_pos - 5, 120, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", car_changes_reported
	  y_pos = y_pos + 10
	  OkButton 510, y_pos, 50, 15
	  CancelButton 560, y_pos, 50, 15
    EndDialog

	Do
		Do
			interview_needed = FALSE
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			If quest_two_changes_reported = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are any changes reported on the Combined AR Form - question 2."
			If car_changes_reported = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are any changes or discrepancies to the known information."

			For each_income = 0 to UBound(INCOME_ARRAY, 2)
				If ButtonPressed = INCOME_ARRAY(btn_const, each_income) Then
					err_msg = "LOOP"
					Call navigate_to_MAXIS_screen("STAT", INCOME_ARRAY(inc_category, each_income))
					EMWriteScreen left(INCOME_ARRAY(owner_name, each_income), 2), 20, 76
					EmWriteScreen right(INCOME_ARRAY(panel_call, each_income), 2), 20, 79
					transmit
				End If
			Next
			If err_msg <> "" AND err_msg <> "LOOP" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

		Loop Until err_msg = ""
		If car_changes_reported = "CHANGES REPORTED" Then interview_needed = TRUE
		If quest_two_changes_reported = "CHANGES REPORTED" Then interview_needed = TRUE

		If interview_needed = TRUE Then
			end_msg = "*** Changes / Differences Reported ***"
			end_msg = end_msg & vbCr & vbCr & "Since the information from the case appears to have potentially changed, we need to complete an interview."
			end_msg = end_msg & vbCr & vbCr & "Providing a full interview for a client with changes provides the best client service and ensures the most accuracy."
			If quest_two_changes_reported = "CHANGES REPORTED" Then end_msg = end_msg & vbCr & vbCr & " - You entered that there are changes reported question two of the Combined AR Form."
			If car_changes_reported = "CHANGES REPORTED" Then end_msg = end_msg & vbCr & vbCr & " - You entered that there are changes reported on the form compared to the information in MAXIS."
			script_end_procedure_with_error_report(end_msg)
		End If

		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE


End If

'IF any other CAF form
	'Earned Income
	'Unearned Income
	'IF MFIP - assets
	'Expenses

If form_type_received <> "Combined Annual Renewal (CAR)" Then
	dlg_len = 130
	earned_grp = 20
	unearned_grp = 20

	UNEA_exists = TRUE
	EARNED_exists = TRUE
	y_pos = 25




	For each_income = 0 to UBound(INCOME_ARRAY, 2)
		If INCOME_ARRAY(category_const, each_income) = "JOBS" Then earned_grp = earned_grp + 10
		If INCOME_ARRAY(category_const, each_income) = "BUSI" Then earned_grp = earned_grp + 10
		If INCOME_ARRAY(category_const, each_income) = "UNEA" Then unearned_grp = unearned_grp + 10

		dlg_len = dlg_len + 10
	Next
	If earned_grp = 20 Then
		earned_grp = 30
		EARNED_exists = FALSE
	End If
	If unearned_grp = 20 Then
		unearned_grp = 30
		UNEA_exists = FALSE
	End If
	' MsgBox earned_grp
	BeginDialog Dialog1, 0, 0, 616, dlg_len, "Income Review"
	  ButtonGroup ButtonPressed
	  GroupBox 10, 10, 600, earned_grp, "Earned Income in MAXIS"
	  If EARNED_exists = FALSE Then
		  Text 20, y_pos, 200, 10, "There is no EARNED INCOME listed on this case in MAXIS."
		  y_pos = y_pos + 10
	  Else
		  For each_income = 0 to UBound(INCOME_ARRAY, 2)
			  If INCOME_ARRAY(category_const, each_income) = "JOBS" Then
				  PushButton 20, y_pos, 45, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 115, 10, "Employer: " & INCOME_ARRAY(employer_const, each_income)
				  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(job_prosp_total_const, each_income)
				  Text 335, y_pos, 45, 10, "Hours:" & INCOME_ARRAY(job_prosp_hours_const, each_income)
				  Text 395, y_pos, 120, 10, "Pay Frequency: " & INCOME_ARRAY(job_freq_const, each_income)
				  y_pos = y_pos + 10
			  End If

			  If INCOME_ARRAY(category_const, each_income) = "BUSI" Then
				  PushButton 20, y_pos, 45, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 115, 10, "Self Employment Type: " & INCOME_ARRAY(busi_type_const, each_income)
				  If INCOME_ARRAY(cash_net_prosp_amount, each_income) <> "0.00" Then
					  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(cash_net_prosp_amount, each_income)
				  Else
					  Text 210, y_pos, 115, 10, "Earnings: $" & INCOME_ARRAY(snap_net_prosp_amount, each_income)
				  End If

				  Text 335, y_pos, 45, 10, "Hours:" & INCOME_ARRAY(reported_hours, each_income)
				  y_pos = y_pos + 10
			  End If
		  Next
	  End If
	  y_pos = y_pos + 10
	  GroupBox 10, y_pos, 600, unearned_grp, "Unearned Income in MAXIS"
	  y_pos = y_pos +15
	  If UNEA_exists = FALSE Then
		  Text 20, y_pos, 200, 10, "There is no UNEARNED INCOME listed on this case in MAXIS."
		  y_pos = y_pos + 10
	  Else
		  For each_income = 0 to UBound(INCOME_ARRAY, 2)
			  If INCOME_ARRAY(category_const, each_income) = "UNEA" Then
				  PushButton 20, y_pos, 50, 10, INCOME_ARRAY(panel_call, each_income), INCOME_ARRAY(btn_const, each_income)
				  Text 85, y_pos, 110, 10, "Source: " & INCOME_ARRAY(unea_type_const, each_income)
				  Text 210, y_pos, 105, 10, "Income Amount: $" & INCOME_ARRAY(unea_amt_const, each_income)
				  y_pos = y_pos + 10
			  End If

		  Next
	  End If
	   y_pos = y_pos + 15
	  Text 20, y_pos, 170, 10, "Does the form report all of these income sources?"
	  DropListBox 190, y_pos - 5, 60, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_report_all_income_sources
	  Text 265, y_pos, 180, 10, "Does the form report any additional income sources?"
	  DropListBox 445, y_pos - 5, 60, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_report_new_income_sources
	  y_pos = y_pos + 20
	  Text 20, y_pos, 380, 10, "Based on this information does it appear any changes or differences are reported, or are there any inconsistencies?"
	  DropListBox 405, y_pos - 5, 180, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", income_changes_apparent
	  y_pos = y_pos + 15
	  ' ButtonGroup ButtonPressed
	    OkButton 510, y_pos, 50, 15
	    CancelButton 560, y_pos, 50, 15
	EndDialog

	Do
		Do
			interview_needed = FALSE
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			If form_report_all_income_sources = "Select One..." Then err_msg = err_msg & vbNewLine & "* Are all of these income sources listed on the form?"
			If form_report_new_income_sources = "Select One..." Then err_msg = err_msg & vbNewLine & "* Are there additional income sources listed on the form?"
			If verifs_report_changes = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are any changes or discrepancies to income."

			For each_income = 0 to UBound(INCOME_ARRAY, 2)
				If ButtonPressed = INCOME_ARRAY(btn_const, each_income) Then
					err_msg = "LOOP"
					Call navigate_to_MAXIS_screen("STAT", INCOME_ARRAY(inc_category, each_income))
					EMWriteScreen left(INCOME_ARRAY(owner_name, each_income), 2), 20, 76
					EmWriteScreen right(INCOME_ARRAY(panel_call, each_income), 2), 20, 79
					transmit
				End If
			Next
			If err_msg <> "" AND err_msg <> "LOOP" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

		Loop Until err_msg = ""
		If income_changes_apparent = "CHANGES REPORTED" Then interview_needed = TRUE

		If form_report_all_income_sources = "No" OR form_report_all_income_sources = "Unsure" Then interview_needed = TRUE
		If form_report_new_income_sources = "Yes" OR form_report_new_income_sources = "Unsure" Then interview_needed = TRUE

		If interview_needed = TRUE Then
			end_msg = "*** Changes / Differences Reported ***"
			end_msg = end_msg & vbCr & vbCr & "Since the information from the case appears to have potentially changed, we need to complete an interview."
			end_msg = end_msg & vbCr & vbCr & "Providing a full interview for a client with changes provides the best client service and ensures the most accuracy."
			If income_changes_apparent = "CHANGES REPORTED" Then end_msg = end_msg & vbCr & vbCr & " - You entered that there are changes reported on the form compared to the information in MAXIS."
			If form_report_all_income_sources = "No" OR form_report_all_income_sources = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form did not report all of the income sources, or that it is unclear."
			If form_report_new_income_sources = "Yes" OR form_report_new_income_sources = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form reported an income source not listed in MAXIS, or that this is unclear."
			script_end_procedure_with_error_report(end_msg)
		End If

		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	dlg_len = 225
	housing_grp_len = 55
	other_grp_len = 45
	shel_exists = FALSE
	dcex_exists = FALSE
	coex_exists = FALSE
	For case_shel = 0 To UBound(SHELTER_ARRAY, 2)
		If SHELTER_ARRAY(shelter_member, case_shel) <> "" Then
			shel_exists = TRUE
			If SHELTER_ARRAY(shel_rent_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_tax_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_room_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_garage_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
			If SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel) <> "" Then
				housing_grp_len = housing_grp_len + 10
				dlg_len = dlg_len + 10
			End If
		End If
	Next

	For case_dcex = 0 to UBound(DCEX_ARRAY, 2)
		If DCEX_ARRAY(dcex_ref_number, case_dcex) <> "" Then
			dcex_exists = TRUE
			other_grp_len = other_grp_len + 10
			dlg_len = dlg_len + 10
		End If
	Next

	For case_coex = 0 to UBound(COEX_ARRAY, 2)
		If COEX_ARRAY(coex_ref_number, case_coex) <> "" Then
			coex_exists = TRUE
			other_grp_len = other_grp_len + 10
			dlg_len = dlg_len + 10
		End If
	Next

	y_pos = 40
	BeginDialog Dialog1, 0, 0, 531, dlg_len, "Expenses Review"
	  Text 205, 5, 80, 10, "--- Review Expenses ---"
	  GroupBox 10, 15, 515, housing_grp_len, "Housing Expenses"
	  Text 20, 30, 50, 10, "Shelter:"
	  If shel_exists = FALSE Then
		  y_pos = 30
		  Text 75, y_pos, 200, 10, "There is no Shelter Expenses Listed in MAXIS"
		  y_pos = y_pos + 10
	  Else
		  For case_shel = 0 To UBound(SHELTER_ARRAY, 2)
		  	  If SHELTER_ARRAY(shelter_member, case_shel) <> "" Then
				  Text 25, y_pos, 125, 10, "Landlord: " & SHELTER_ARRAY(shel_paid_to, case_shel)
				  Text 395, y_pos, 100, 10, "Listed under: MEMB " & SHELTER_ARRAY(shelter_member, case_shel)
					If SHELTER_ARRAY(shel_rent_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Rent: $" & SHELTER_ARRAY(shel_rent_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_rent_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Lot Rent: $" & SHELTER_ARRAY(shel_lot_rent_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_lot_rent_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Mortgage: $" & SHELTER_ARRAY(shel_mortgage_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_mortgage_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Insurance: $" & SHELTER_ARRAY(shel_insurance_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_insurance_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_tax_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Tax: $" & SHELTER_ARRAY(shel_tax_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_tax_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_room_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Room: $" & SHELTER_ARRAY(shel_room_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_room_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_garage_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Garage: $" & SHELTER_ARRAY(shel_garage_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_garage_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
					If SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel) <> "" Then
						Text 155, y_pos, 70, 10, "Subsidy: $" & SHELTER_ARRAY(shel_subsidy_prosp_amt, case_shel)
						Text 240, y_pos, 140, 10, "Verif: " & SHELTER_ARRAY(shel_subsidy_prosp_verif, case_shel)
						y_pos = y_pos + 10
					End If
			  End If
		  Next
	  End If
	  y_pos = y_pos + 5
	  Text 20, y_pos, 35, 10, "Utilities:"
	  y_pos = y_pos + 10
	  Text 30, y_pos, 50, 10, "Heat/AC: " & Heat_Air_YN
	  Text 85, y_pos, 70, 10, "Amount - $" & Heat_Air_Amount
	  Text 205, y_pos, 50, 10, "Electric: " & Electric_YN
	  Text 260, y_pos, 70, 10, "Amount - $" & Electric_Amount
	  Text 370, y_pos, 50, 10, "Phone: " & Phone_YN
	  Text 425, y_pos, 70, 10, "Amount - $" & Phone_Amount
	  y_pos = y_pos + 20
	  Text 15, y_pos + 5, 135, 10, "Does the form list this housing expense?"
	  DropListBox 150, y_pos, 45, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_list_housing_expense
	  Text 205, y_pos + 5, 110, 10, "Does the form list these utilities?"
	  DropListBox 310, y_pos, 45, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_list_utilities
	  Text 370, y_pos + 5, 115, 10, "Does the form list any additional?"
	  DropListBox 475, y_pos, 45, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_list_more_housing_expense
	  y_pos = y_pos + 20
	  Text 20, y_pos + 5, 235, 10, "Does there appear to be any changes to any of the housing expenses?"
	  DropListBox 260, y_pos, 260, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", housing_expense_change_apparent
	  y_pos = y_pos + 20
	  GroupBox 10, y_pos, 515, other_grp_len, "Other Expenses"
	  y_pos = y_pos + 15
	  Text 20, y_pos, 90, 10, "Dependent Care Expense:"
	  If dcex_exists = FALSE Then
		  Text 120, y_pos, 200, 10, "There are no DEPENDENT CARE EXPENSES listed in MAXIS."
		  y_pos = y_pos + 10
	  Else
	  	  y_pos = y_pos + 10
		  For case_dcex = 0 to UBound(DCEX_ARRAY, 2)
	  		  If DCEX_ARRAY(dcex_ref_number, case_dcex) <> "" Then
				  Text 25, y_pos, 110, 10, "Provider: " & DCEX_ARRAY(provider_const, case_dcex)
				  Text 140, y_pos, 105, 10, "For: " & DCEX_ARRAY(child_list_const, case_dcex)
				  Text 255, y_pos, 65, 10, "Amount: $" & DCEX_ARRAY(total_amt_const, case_dcex)
				  Text 330, y_pos, 90, 10, "Verif: " & DCEX_ARRAY(dcex_verif_const, case_dcex)
				  Text 425, y_pos, 95, 10, "Reason: " & DCEX_ARRAY(dcex_reason_const, case_dcex)
				  y_pos = y_pos + 10
			  End If
		  Next
	  End If
	  y_pos = y_pos + 5
	  Text 20, y_pos, 90, 10, "Court Ordered Expense:"
	  If coex_exists = FALSE Then
		  Text 120, y_pos, 200, 10, "There are no COURT ORDERED EXPENSES listed in MAXIS."
		  y_pos = y_pos + 10
	  Else
		  y_pos = y_pos + 10
		  For case_coex = 0 to UBound(COEX_ARRAY, 2)
		  	  If COEX_ARRAY(coex_ref_number, case_coex) <> "" Then
				  Text 25, y_pos, 70, 10, "Paid by: MEMB " & COEX_ARRAY(coex_ref_number, case_coex)
				  Text 100, y_pos, 80, 10, "Support: Amt - $" & COEX_ARRAY(support_amount_const, case_coex)
				  Text 185, y_pos, 80, 10, "Alimony: Amt - $" & COEX_ARRAY(alimony_amount_const, case_coex)
				  Text 270, y_pos, 80, 10, "Tax Dep: Amt - $" & COEX_ARRAY(tax_dep_amount_const, case_coex)
				  Text 355, y_pos, 80, 10, "Other: Amt - $" & COEX_ARRAY(other_amount_const, case_coex)
				  y_pos = y_pos + 10
			  End If
		  Next
	  End If
	  y_pos = y_pos + 15
	  Text 20, y_pos, 155, 10, "Does the form list all of these other expenses?"
	  DropListBox 175, y_pos - 5, 60, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_list_other_expenses
	  Text 270, y_pos, 130, 10, "Does the form list any other expenses?"
	  DropListBox 400, y_pos - 5, 60, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Unsure", form_list_more_other_expenses
	  y_pos = y_pos + 20
	  Text 45, y_pos, 210, 10, "Does there appear to be any changes to these other expenses?"
	  DropListBox 260, y_pos - 5, 260, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change", other_expense_change_apparent
	  y_pos = y_pos + 15
	  ButtonGroup ButtonPressed
	    CancelButton 475, y_pos, 50, 15
	    OkButton 420, y_pos, 50, 15
	EndDialog

	Do
		Do
			interview_needed = FALSE
			err_msg = ""

			dialog Dialog1
			cancel_confirmation


			If form_list_housing_expense = "Select..." Then err_msg = err_msg & vbNewLine & "* Are all of the housing expenses listed on the form?"
			If form_list_utilities = "Select..." Then err_msg = err_msg & vbNewLine & "* Are all of the utilities listed on the form?"
			If form_list_more_housing_expense = "Select..." Then err_msg = err_msg & vbNewLine & "* Are there more shelter expenses listed on the form?"
			If form_list_other_expenses = "Select One..." Then err_msg = err_msg & vbNewLine & "* Are all of these other expenses listed on the form?"
			If form_list_more_other_expenses = "Select One..." Then err_msg = err_msg & vbNewLine & "* Are there more expenses listed on the form?"

			If housing_expense_change_apparent = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are any changes or discrepenies in the shelter expenses."
			If other_expense_change_apparent = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are any changes or discrepencies in any of the expenses."

			If err_msg <> "" AND err_msg <> "LOOP" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

		Loop Until err_msg = ""
		If housing_expense_change_apparent = "CHANGES REPORTED" Then interview_needed = TRUE
		If other_expense_change_apparent = "CHANGES REPORTED" Then interview_needed = TRUE

		If form_list_housing_expense = "No" OR form_list_housing_expense = "Unsure" Then interview_needed = TRUE
		If form_list_utilities = "No" OR form_list_utilities = "Unsure" Then interview_needed = TRUE
		If form_list_more_housing_expense = "Yes" OR form_list_more_housing_expense = "Unsure" Then interview_needed = TRUE
		If form_list_other_expenses = "No" OR form_list_other_expenses = "Unsure" Then interview_needed = TRUE
		If form_list_more_other_expenses = "Yes" OR form_list_more_other_expenses = "Unsure" Then interview_needed = TRUE

		If interview_needed = TRUE Then
			end_msg = "*** Changes / Differences Reported ***"
			end_msg = end_msg & vbCr & vbCr & "Since the information from the case appears to have potentially changed, we need to complete an interview."
			end_msg = end_msg & vbCr & vbCr & "Providing a full interview for a client with changes provides the best client service and ensures the most accuracy."

			If housing_expense_change_apparent = "CHANGES REPORTED" Then end_msg = end_msg & vbCr & vbCr & " - You entered that there are changes in shelter expense reported on the form compared to the information in MAXIS."
			If other_expense_change_apparent = "CHANGES REPORTED" Then end_msg = end_msg & vbCr & vbCr & " - You entered that there are changes in other expense reported on the form compared to the information in MAXIS."

			If form_list_housing_expense = "No" OR form_list_housing_expense = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form did not list all of the shelter expenses, or that they were unclear."
			If form_list_utilities = "No" OR form_list_utilities = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form did not list all of the utilities expenses, or that they were unclear."
			If form_list_more_housing_expense = "Yes" OR form_list_more_housing_expense = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form reported a shelter expense not listed in MAXIS, or that they were unclear."
			If form_list_other_expenses = "No" OR form_list_other_expenses = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form did not report all of the other expenses, or that they were unclear."
			If form_list_more_other_expenses = "Yes" OR form_list_more_other_expenses = "Unsure" Then end_msg = end_msg & vbCr & vbCr & " - You indicated that the form reported an other expense not listed in MAXIS, or that they were unclear."

			script_end_procedure_with_error_report(end_msg)
		End If

		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

End If
'ALL - check verifications received with the form and ask if there is anything else
'Offer to send a 'second review' to KN

BeginDialog Dialog1, 0, 0, 426, 250, "Final Review"
  DropListBox 300, 125, 120, 45, "Select One..."+chr(9)+"CHANGES REPORTED"+chr(9)+"Appears No Change"+chr(9)+"No docs Attached", verifs_report_changes
  EditBox 5, 155, 415, 15, waived_explanation
  DropListBox 120, 205, 265, 45, "No - I was able to determine if we should waive the interivew"+chr(9)+"Yes - please send this case for additional review", request_additional_support
  ButtonGroup ButtonPressed
    OkButton 320, 230, 50, 15
    CancelButton 370, 230, 50, 15
  Text 100, 35, 235, 15, "Review any documents for information that is new, different, or indicates a change that was not clear in the information added to the form."
  Text 25, 65, 100, 10, "Common issues to check for:"
  Text 40, 80, 215, 10, "- Irregular income (bonuses, overtime, inconsitent sources, etc)."
  Text 40, 90, 215, 10, "- New Expense sources"
  Text 40, 100, 215, 10, "- Unknown household members/people"
  Text 40, 110, 215, 10, "- References to schools/disability/jobs that are not known."
  Text 20, 195, 370, 10, "Are you unsure about any of this information? Do you want someone to take a second look at the case and form?"
  GroupBox 5, 180, 415, 45, "Need Support"
  Text 20, 130, 280, 10, "Any changes or unclear information listed in the verifications/additional documents?"
  Text 5, 145, 195, 10, "Explain how interview can be waived (for the CASE:NOTE):"
  Text 95, 15, 245, 10, "--- Check ECF for any additional documentation or verifications received. ---"
EndDialog

Do
	Do
		interview_needed = FALSE
		err_msg = ""

		dialog Dialog1
		cancel_confirmation

		If verifs_report_changes = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if any additional documents suggest a change or discrepancy."

		If err_msg <> "" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

	Loop Until err_msg = ""
	If verifs_report_changes = "CHANGES REPORTED" Then interview_needed = TRUE

	If interview_needed = TRUE Then
		end_msg = "*** Changes / Differences Reported ***"
		end_msg = end_msg & vbCr & vbCr & "Since the information from the case appears to have potentially changed, we need to complete an interview."
		end_msg = end_msg & vbCr & vbCr & "Providing a full interview for a client with changes provides the best client service and ensures the most accuracy."
		script_end_procedure_with_error_report(end_msg)
	End If

	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

If request_additional_support = "Yes - please send this case for additional review" Then
	Call find_user_name(worker_name)

	list_the_progs = ""
	If snap_er_checkbox = checked Then list_the_progs = list_the_progs & "SNAP, "
	If mfip_er_checkbox = checked Then list_the_progs = list_the_progs & "MFIP, "
	If grh_er_checkbox = checked Then list_the_progs = list_the_progs & "GRH, "
	If ga_er_checkbox = checked Then list_the_progs = list_the_progs & "GA, "
	If msa_er_checkbox = checked Then list_the_progs = list_the_progs & "MSA, "
	If right(list_the_progs, 2) = ", " Then list_the_progs = left(list_the_progs, len(list_the_progs) - 2)

	BeginDialog Dialog1, 0, 0, 401, 185, "Send Request to Review to QI"
	  EditBox 10, 140, 385, 15, questions_for_qi
	  EditBox 70, 165, 110, 15, worker_name
	  ButtonGroup ButtonPressed
	    PushButton 190, 165, 80, 15, "Send EMAIL to QI - KN", send_email_btn
	    PushButton 270, 165, 125, 15, "Cancel EMAIL - Waive ER Interview", cancel_email_btn
	  Text 10, 10, 395, 10, "You have requested to send an email request to QI Knowledge Now for a review of possible ER Interview to be waived."
	  GroupBox 75, 25, 220, 90, "Case Information"
	  Text 90, 40, 85, 10, "Case Number: " & MAXIS_case_number
	  Text 90, 60, 150, 10, "Review Month: " & recert_month & " " & recert_year
	  Text 90, 80, 175, 10, "Programs to REVW: " & list_the_progs
	  Text 90, 100, 85, 10, "CAF Date: " & caf_date
	  Text 10, 125, 115, 10, "List any specific questions here:"
	  Text 10, 170, 60, 10, "Sign your Email:"
	EndDialog

	Do
		Do
			interview_needed = FALSE
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			If err_msg <> "" Then MsgBox "**** NOTICE *****" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

		Loop Until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	questions_for_qi = trim(questions_for_qi)

	If ButtonPressed = send_email_btn Then
		email_subject = "PLEASE REVIEW Case #" & MAXIS_case_number & " for Waived ER Interview"
		email_body = "I have reviewed the case for a potential waived ER interview, but would like a second look."
		email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr
		email_body = email_body & "REVW Month: " & recert_month & " " & recert_year & vbCr
		email_body = email_body & "Programs to REVW: " & list_the_progs & vbCr
		email_body = email_body & "CAF received on: " & caf_date & vbCr & vbCR
		If questions_for_qi <> "" Then email_body = email_body & "Specific questions about the case: " & questions_for_qi & vbCr & vbCR

		email_body = email_body & vbCr & "Thank you, " & vbCr & worker_name

		Call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", TRUE)		'Send the Email

		end_msg = "The email has been sent to QI." & vbCR & vbCR & "The script will now end as QI will need time to review and respond."

		script_end_procedure_with_error_report(end_msg)
	End If
End If

'Case appears to meet criteria to have the interview waived. PLEASE CONFIRM.
confirm_interview_waived = MsgBox("Based on your responses to all of the questions, it appears this case is eligible to have the interview waived." & vbCr & vbCr & "To complete the process steps in making this assesment, the case should be updated with:" & vbCr & "  - A clear CASE:NOTE of this determination." & vbCr & "  - Update the panel STAT:REVW with today's date for the interview date." & vbCr & vbCr & "This script will take these steps after this message." & vbCr & vbCr & "Press OK to confirm that this ER Interview should be waived so the script can continue.", vbImportant + vbOkCancel, "Confirm the Interview should be Waived")
If confirm_interview_waived = vbCancel Then
	end_msg = "Though all information previously suggested the interview should be waived, you did not confirm this assesment." & vbCr & vbCr & "The script has NOT updated REVW or entered a CASE:NOTE."
	script_end_procedure_with_error_report(end_msg)
End If

Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "REVW")

EMWriteScreen "I", 7, 40
EMWriteScreen "I", 7, 60
Call create_mainframe_friendly_date(caf_date, 13, 37, "YY")
Call create_mainframe_friendly_date(date, 15, 37, "YY")

' explanation = "This case has: "
' If UNEA_exists = FALSE Then
' 	explanation = explanation & "no unearned income, "
' If EARNED_exists = FALSE Then
' 	explanation = explanation & "no earned income, "
' If shel_exists = FALSE Then
' 	explanation = explanation & "no shelter expense, "
' If dcex_exists = FALSE Then explanation = explanation & "no dependent care expense, "
' If coex_exists = FALSE Then explanation = explanation & "no court ordered expenses, "

Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("ER Interview Waived for "& MAXIS_footer_month & "/" & MAXIS_footer_year &" Recertification")
Call write_variable_in_CASE_NOTE("Interview requirement waived after CAF and Case review completed on " & date &".")
Call write_variable_in_CASE_NOTE("Case is waived: " & waived_explanation)
Call write_variable_in_CASE_NOTE("REVW updated with an interview date of " & date & " as that is the day the case was reviewed and the interview determined to be waived.")
Call write_variable_in_CASE_NOTE("Ability to waive recertification interviews granted from FNS and Emergency Order 20-12.")
Call write_variable_in_CASE_NOTE("Full information about the case and information used to determine eligibility and budgeting in following case note.")
Call write_variable_in_CASE_NOTE("The interview may still be needed if information is unable to be determined without clarification from the client or if the client requests an interview is completed.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

end_msg = "The REVW panel has been updated and CASE:NOTE entered." & vbCr & vbCr & "--- ADDITIONAL CASE PROCESSING NEEDED ---" & vbCr & vbCr & "This note is NOT sufficient for a complete ER CASE:NOTE. You must complete the processing of the ER in STAT and send any necessary verification requests and create a complete CASE:NOTE of the case situation and ER processing. The CAF script is availbale and has functionality to allow the interview information to be waived in the script." & vbCr & vbCr & "Failure to complete processing and a CASE:NOTE could cause an error and inaccurate payments for this case."
script_end_procedure_with_error_report(end_msg)
