'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ADDRESS AND SHELTER.vbs"
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



'MAP of PROCESS
'Create FUNCTIONS - make them able to be used gloally
	'Read PERSONS who are ACTIVE SNAP/Cash
	'READ and WRITE ADDR - verif and not

function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, verif_received)
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
		' verif_received 'add functionality to change how this is updated based on if we have verif or not.
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


	'READ and WRITE SHEL - verif and not - handle for MEMBERS

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


	'READ and WRITE HEST

function access_HEST_panel(access_type, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "HEST")

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
        retro_electric_yn = replace(retro_electric_yn, "_", "")
        retro_electric_units = replace(retro_electric_units, "_", "")
        retro_electric_amt = trim(retro_electric_amt)
        retro_phone_yn = replace(retro_phone_yn, "_", "")
        retro_phone_units = replace(retro_phone_units, "_", "")
        retro_phone_amt = trim(retro_phone_amt)

        prosp_heat_ac_yn = replace(prosp_heat_ac_yn, "_", "")
        prosp_heat_ac_units = replace(prosp_heat_ac_units, "_", "")
        prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
        If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
        prosp_electric_yn = replace(prosp_electric_yn, "_", "")
        prosp_electric_units = replace(prosp_electric_units, "_", "")
        prosp_electric_amt = trim(prosp_electric_amt)
        If prosp_electric_amt = "" Then prosp_electric_amt = 0
        prosp_phone_yn = replace(prosp_phone_yn, "_", "")
        prosp_phone_units = replace(prosp_phone_units, "_", "")
        prosp_phone_amt = trim(prosp_phone_amt)
        If prosp_phone_amt = "" Then prosp_phone_amt = 0

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
end function

	'Dialog to gather address
	'Dialog to gather SHEL and HEST
	'Output for update
	'transfer functionality
	'CASE:NOTE new address and new shel

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

'THIS script
	'Initial Dialog
		'Case Number
		'Worker Sig
		'NEW ADDR - Y/N
		'NEW SHEL - Y/N

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to Bluezone & grabbing case number and footer year/month
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 356, 115, "Dialog"
  EditBox 65, 10, 50, 15, MAXIS_case_number
  CheckBox 135, 5, 210, 10, "Check here if the change was previously reported and panels", shelter_change_previously_reported
  Text 145, 15, 155, 10, "updated and this action is only for verifiication."
  DropListBox 200, 35, 60, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", new_addr_yn
  EditBox 300, 35, 50, 15, addr_change_date
  DropListBox 200, 55, 60, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", new_shel_yn
  EditBox 300, 55, 50, 15, shel_change_date
  EditBox 200, 75, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 245, 95, 50, 15
    CancelButton 300, 95, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 40, 40, 160, 10, "Is there a change to the household's address?"
  Text 270, 40, 25, 10, "on date"
  Text 10, 60, 185, 10, "Is there a change to the household's shelter expenses?"
  Text 270, 60, 25, 10, "on date"
  Text 135, 80, 65, 10, "Worker Signature:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1																'calls up dialog for worker to enter case number and applicable month and year.	 Script will 'loop'
		cancel_without_confirmation

		call validate_MAXIS_case_number("*", err_msg)
		If new_addr_yn = "Select" AND new_shel_yn = "Select" AND shelter_change_previously_reported = unchecked Then err_msg = err_sg & vbNewLine & "* Indicate if there is a change in Address or Shelter Expenses."

		If new_addr_yn = "Yes" Then
			If IsDate(addr_change_date) = FALSE Then err_msg = err_sg & vbNewLine & "* Since you have indicated that the Address has changed, please enter the date of change as a valid date."
		End If

		If new_shel_yn = "Yes" Then
			If IsDate(shel_change_date) = FALSE Then err_msg = err_sg & vbNewLine & "* Since you have indicated that the Shelter Expense has changed, please enter the date of change as a valid date."
		End If

		If IsDate(shel_change_date) = TRUE AND IsDate(addr_change_date) = TRUE Then
			If shel_change_date <> addr_change_date Then err_msg = err_msg & vbNewLine & "* The date of change for Address and Shelter Expense should match."
		End If

		If err_msg = "" Then MsgBox "****** NOTICE *******" & vbNewLine & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg

	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

If new_addr_yn = "Yes" Then
	new_shel_yn = "Yes"
	shel_change_date = addr_change_date
End If

If shelter_change_previously_reported = checked Then
	'maybe read a CASE:;NOTE from a previous run to fill the information to update
	call navigate_to_MAXIS_screen("CASE", "NOTE")

	note_role = 5
	Do
		EMReadScreen note_date, 8, note_row, 6
		EMReadScreen note_title, 55, note_row, 25

		If note_date = "        " Then Exit Do
		note_title = trim(note_title)

		If note_tite = "" Then 				'look for a note created by this script to gather previously known information
			'Get the update date - determine if ADDR AND SHEL or SHEL ONLY
		End If

	Loop until
End If

call convert_date_into_MAXIS_footer_month(shel_change_date, MAXIS_footer_month, MAXIS_footer_year)
call back_to_SELF

Do
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "SUMM", is_this_priv)
	If is_this_priv = TRUE Then script_end_procedure("This case is priivileged and you do not appear to have access. If you need access to process this case, please request access first and run this script agin.")
	EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"
Call MAXIS_footer_month_confirmation

	'Figure out if we can update or not (inquiry)
continue_in_inquiry = ""
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If


	'Read SHEL HEST and  ADDR
	'IF NEW ADDR - ADDR DIALOG
	'IF NEW SHEL - SHEL DIALOG
	'If able to update - update, otherwise - output
	'transfer if neeeded and able

'Client Contact - maybe have a whole list of buttons on the side for different functionality
	'Add button to indicate a change
'Changes Reported 
'Docs Received
'Interview Script
