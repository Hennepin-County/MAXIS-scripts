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

' function access_JOBS_panel(access_type, job_member, job_verif, job_employer, job_type, job_pay_amount, job_prosp_total, job_prosp_hours, job_frequency, job_update_date, job_start_date, job_end_date, panel_ref_numb, hourly_wage, retrospective_total, retrospective_hours, fs_pic_pay_frequency, fs_pic_average_hours, fs_pic_average_pay, fs_pic_monthly_prospective, grh_pic_pay_frequency, grh_pic_average_pay, grh_pic_monthly_prospective, jobs_subsidy_code)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen job_member, 2, 4, 33
'         EMReadScreen job_type, 1, 5, 34
'         EMReadScreen jobs_subsidy_code, 2, 5, 74
'         EMReadScreen job_verif, 1, 6, 34
'         EMReadScreen hourly_wage, 6, 6, 75
'         EMReadScreen employer_name, 30, 7, 42
'         EMReadScreen income_start_date, 8, 9, 35
'         EMReadScreen income_end_date, 8, 9, 49
'         EMReadScreen retrospective_total, 8, 17, 38
'         EMReadScreen retrospective_hours, 3, 18, 43
'
'         For jobs_row = 16 to 12 Step -1
'             EMReadScreen paycheck_amount, 8, jobs_row, 67
'             If paycheck_amount <> "________" Then
'                 job_pay_amount = trim(paycheck_amount)
'                 Exit For
'             End If
'         Next
'         EMReadScreen job_prosp_total, 8, 17, 67
'         EMReadScreen job_frequency, 1, 18, 35
'         EMReadScreen job_prosp_hours, 3, 18, 72
'         EMReadScreen last_updated, 8, 21, 55
'         ' MsgBox "Line 817" & vbNewLine & last_updated
'
'         EMWriteScreen "X", 19, 38           'opening the FS PIC
'         transmit
'
'         EMReadScreen fs_pic_pay_frequency, 1, 5, 64
'         EMReadScreen fs_pic_average_hours, 6, 16, 51
'         EMReadScreen fs_pic_average_pay, 8, 17, 56
'         EMReadScreen fs_pic_monthly_prospective, 8, 18, 56
'
'         PF3                                 'closing the FS PIC
'
'         EMWriteScreen "X", 19, 71           'opening the GRH PIC
'         transmit
'
'         EMReadScreen grh_pic_pay_frequency, 1, 3, 63
'         EMReadScreen grh_pic_average_pay, 8, 16, 65
'         EMReadScreen grh_pic_monthly_prospective, 8, 17, 65
'
'         PF3                                 'closing the GRH PIC
'
'         If jobs_subsidy_code = "01" Then jobs_subsidy_code = "01 - Subsidized Public Secotr Employer"
'         If jobs_subsidy_code = "02" Then jobs_subsidy_code = "02 - Subsidized Private Sector Employer"
'         If jobs_subsidy_code = "03" Then jobs_subsidy_code = "03 - On-the-Job-Training"
'         If jobs_subsidy_code = "04" Then jobs_subsidy_code = "04 - Americorps"
'         If jobs_subsidy_code = "__" Then jobs_subsidy_code = "None"
'
'         hourly_wage = trim(hourly_wage)
'         retrospective_total = trim(retrospective_total)
'         retrospective_hours = trim(retrospective_hours)
'
'         job_employer = replace(employer_name, "_", "")
'         If job_verif = "1" Then job_verif = "1 - Pay Stubs"
'         If job_verif = "2" Then job_verif = "2 - Empl Stmt"
'         If job_verif = "3" Then job_verif = "3 - Coltrl Stmt"
'         If job_verif = "4" Then job_verif = "4 - Other Doc"
'         If job_verif = "5" Then job_verif = "5 - Pend Out State"
'         If job_verif = "N" Then job_verif = "N - No Verif Prvd"
'         If job_verif = "?" Then job_verif = "? - Delayed Verif"
'
'         If job_type = "J" Then job_type = "J - WIOA"
'         If job_type = "W" Then job_type = "W - Wages"
'         If job_type = "E" Then job_type = "E - EITC"
'         If job_type = "G" Then job_type = "G - Experience Works"
'         If job_type = "F" Then job_type = "F - Fed Work Study"
'         If job_type = "S" Then job_type = "S - State Work Study"
'         If job_type = "O" Then job_type = "O - Other"
'         If job_type = "C" Then job_type = "C - Contract Income"
'         If job_type = "T" Then job_type = "T - Training Prog"
'         If job_type = "P" Then job_type = "P - Service Prog"
'         If job_type = "R" Then job_type = "R - Rehab Prog"
'
'         job_prosp_total = trim(job_prosp_total)
'         job_prosp_hours = trim(job_prosp_hours)
'         If job_frequency = "1" Then job_frequency = "1 - Monthly"
'         If job_frequency = "2" Then job_frequency = "2 - Semi Monthly"
'         If job_frequency = "3" Then job_frequency = "3 - Biweekly"
'         If job_frequency = "4" Then job_frequency = "4 -  Weekly"
'         If job_frequency = "5" Then job_frequency = "5 - Other"
'
'         job_update_date = replace(last_updated, " ", "/")
'         ' MsgBox "Line 849" & vbNewLine & job_update_date
'         job_start_date = replace(income_start_date, " ", "/")
'         job_end_date = replace(income_end_date, " ", "/")
'         if job_end_date = "__/__/__" then job_end_date = ""
'
'         If fs_pic_pay_frequency = "1" Then fs_pic_pay_frequency = "1 - Monthly"
'         If fs_pic_pay_frequency = "2" Then fs_pic_pay_frequency = "2 - Semi Monthly"
'         If fs_pic_pay_frequency = "3" Then fs_pic_pay_frequency = "3 - Biweekly"
'         If fs_pic_pay_frequency = "4" Then fs_pic_pay_frequency = "4 -  Weekly"
'         If fs_pic_pay_frequency = "5" Then fs_pic_pay_frequency = "5 - Other"
'         If fs_pic_pay_frequency = "_" Then fs_pic_pay_frequency = ""
'
'         fs_pic_average_hours = trim(fs_pic_average_hours)
'         fs_pic_average_pay = trim(fs_pic_average_pay)
'         fs_pic_monthly_prospective = trim(fs_pic_monthly_prospective)
'
'         If grh_pic_pay_frequency = "1" Then grh_pic_pay_frequency = "1 - Monthly"
'         If grh_pic_pay_frequency = "2" Then grh_pic_pay_frequency = "2 - Semi Monthly"
'         If grh_pic_pay_frequency = "3" Then grh_pic_pay_frequency = "3 - Biweekly"
'         If grh_pic_pay_frequency = "4" Then grh_pic_pay_frequency = "4 -  Weekly"
'         If grh_pic_pay_frequency = "5" Then grh_pic_pay_frequency = "5 - Other"
'         If grh_pic_pay_frequency = "_" Then grh_pic_pay_frequency = ""
'
'         grh_pic_average_pay = trim(grh_pic_average_pay)
'         grh_pic_monthly_prospective = trim(grh_pic_monthly_prospective)
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_BUSI_panel(access_type, busi_member, busi_type, income_start_date, income_end_date, cash_net_prosp_amount, cash_net_retro_amount, cash_retro_total_income, cash_retro_expenses, cash_prosp_total_income, cash_prosp_expenses, cash_income_verif, cash_expense_verif, snap_net_prosp_amount, snap_net_retro_amount, snap_retro_total_income, snap_retro_expenses, snap_prosp_total_income, snap_prosp_expenses, snap_income_verif, snap_expense_verif, hc_method_a_net_prosp_amount, hc_method_b_net_prosp_amount, hc_method_a_total_income, hc_method_a_expenses, hc_method_a_income_verif, hc_method_a_expense_verif, hc_method_b_total_income, hc_method_b_expenses, hc_method_b_income_verif, hc_method_b_expense_verif, SE_method, SE_method_date, reported_hours, minimum_wage_hours, update_date, panel_ref_numb)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen busi_member, 2, 4, 33
'         EMReadScreen type_of_income, 2, 5, 37
'         EMReadScreen income_start_date, 8, 5, 55
'         EMReadScreen income_end_date, 8, 5, 72
'         EMReadScreen cash_net_prosp_amount, 8, 8, 69
'         EMReadScreen cash_net_retro_amount, 8, 8, 55
'         EMReadScreen snap_net_prosp_amount, 8, 10, 69
'         EMReadScreen snap_net_retro_amount, 8, 10, 55
'         EMReadScreen hc_method_a_net_prosp_amount, 8, 11, 69
'         EMReadScreen hc_method_b_net_prosp_amount, 8, 12, 69
'         EMReadScreen reported_hours, 3, 13, 74
'         EMReadScreen minimum_wage_hours, 3, 14, 74
'         EMReadScreen update_date, 8, 21, 55
'
'         EMReadScreen SE_method, 2, 16, 53
'         EMReadScreen SE_method_date, 8, 16, 63
'
'         ' MsgBox "Line 870" & vbNewLine & update_date
'         EMWriteScreen "X", 6, 26
'         transmit
'         EMReadScreen cash_retro_total_income, 8, 9, 43
'         EMReadScreen cash_retro_expenses, 8, 15, 43
'         EMReadScreen cash_prosp_total_income, 8, 9, 59
'         EMReadScreen cash_prosp_expenses, 8, 15, 59
'         EMReadScreen cash_income_verif, 1, 9, 73
'         EMReadScreen cash_expense_verif, 1, 15, 73
'
'         EMReadScreen snap_retro_total_income, 8,  11, 43
'         EMReadScreen snap_retro_expenses, 8, 17, 43
'         EMReadScreen snap_prosp_total_income, 8,  11, 59
'         EMReadScreen snap_prosp_expenses, 8, 17, 59
'         EMReadScreen snap_income_verif, 1, 11, 73
'         EMReadScreen snap_expense_verif, 1, 17, 73
'
'         EMReadScreen hc_method_a_total_income, 8, 12, 59
'         EMReadScreen hc_method_a_income_verif, 1, 12, 73
'         EMReadScreen hc_method_a_expenses, 8, 18, 59
'         EMReadScreen hc_method_a_expense_verif, 1, 18, 73
'
'         EMReadScreen hc_method_b_total_income, 8, 13, 59
'         EMReadScreen hc_method_b_income_verif, 1, 13, 73
'         EMReadScreen hc_method_b_expenses, 8, 19, 59
'         EMReadScreen hc_method_b_expense_verif, 1, 19, 73
'
'         PF3
'
'         If type_of_income = "01" Then busi_type = "01 - Farming"
'         If type_of_income = "02" Then busi_type = "02 - Real Estate"
'         If type_of_income = "03" Then busi_type = "03 - Home Product Sales"
'         If type_of_income = "04" Then busi_type = "04 - Other Sales"
'         If type_of_income = "05" Then busi_type = "05 - Personal Services"
'         If type_of_income = "06" Then busi_type = "06 - Paper Route"
'         If type_of_income = "07" Then busi_type = "07 - In Home Daycare"
'         If type_of_income = "08" Then busi_type = "08 - Rental Income"
'         If type_of_income = "09" Then busi_type = "09 - Other"
'
'         income_start_date = replace(income_start_date, " ", "/")
'         income_end_date = replace(income_end_date, " ", "/")
'         If income_end_date = "__/__/__" Then income_end_date = ""
'         ' cash_net_prosp_amount = trim(cash_net_prosp_amount)
'         ' snap_net_prosp_amount = trim(snap_net_prosp_amount)
'         ' hc_method_a_net_prosp_amount = trim(hc_method_a_net_prosp_amount)
'         ' hc_method_b_net_prosp_amount = trim(hc_method_b_net_prosp_amount)
'         ' reported_hours = trim(reported_hours)
'         ' minimum_wage_hours = trim(minimum_wage_hours)
'         update_date = replace(update_date, " ", "/")
'
'         If SE_method = "01" Then SE_method = "01 - 50% Gross Inc"
'         If SE_method = "02" Then SE_method = "02 - Tax Forms"
'
'         SE_method_date = replace(SE_method_date, " ", "/")
'         If SE_method_date = "__/__/__" Then SE_method_date = ""
'
'         ' MsgBox "Line 898" & vbNewLine & update_date
'         If cash_income_verif = "1" Then cash_verif = "1 - Tax Returns"
'         If cash_income_verif = "2" Then cash_verif = "2 - Receipts"
'         If cash_income_verif = "3" Then cash_verif = "3 - Busi Records"
'         If cash_income_verif = "6" Then cash_verif = "6 - Other Doc"
'         If cash_income_verif = "N" Then cash_verif = "N - No Verif Prvd"
'         If cash_income_verif = "?" Then cash_verif = "? - Delayed Verif"
'
'         If snap_income_verif = "1" Then snap_verif = "1 - Tax Returns"
'         If snap_income_verif = "2" Then snap_verif = "2 - Receipts"
'         If snap_income_verif = "3" Then snap_verif = "3 - Busi Records"
'         If snap_income_verif = "4" Then snap_verif = "4 - Pend Out State"
'         If snap_income_verif = "6" Then snap_verif = "6 - Other Doc"
'         If snap_income_verif = "N" Then snap_verif = "N - No Verif Prvd"
'         If snap_income_verif = "?" Then snap_verif = "? - Delayed Verif"
'
'         If hc_method_b_income_verif = "1" Then hc_verif = "1 - Tax Returns"
'         If hc_method_b_income_verif = "2" Then hc_verif = "2 - Receipts"
'         If hc_method_b_income_verif = "3" Then hc_verif = "3 - Busi Records"
'         If hc_method_b_income_verif = "6" Then hc_verif = "6 - Other Doc"
'         If hc_method_b_income_verif = "N" Then hc_verif = "N - No Verif Prvd"
'         If hc_method_b_income_verif = "?" Then hc_verif = "? - Delayed Verif"
'
'         cash_retro_total_income = replace(cash_retro_total_income, "_", " ")
'         ' cash_retro_total_income = trim(cash_retro_total_income)
'
'         cash_retro_expenses = replace(cash_retro_expenses, "_", " ")
'         ' cash_retro_expenses = trim(cash_retro_expenses)
'
'         cash_prosp_total_income = replace(cash_prosp_total_income, "_", " ")
'         ' cash_prosp_total_income = trim(cash_prosp_total_income)
'
'         cash_prosp_expenses = replace(cash_prosp_expenses, "_", " ")
'         ' cash_prosp_expenses = trim(cash_prosp_expenses)
'
'         If cash_income_verif = "1" Then cash_income_verif = "1 - Tax Returns"
'         If cash_income_verif = "2" Then cash_income_verif = "2 - Receipts"
'         If cash_income_verif = "3" Then cash_income_verif = "3 - Busi Records"
'         If cash_income_verif = "6" Then cash_income_verif = "6 - Other Doc"
'         If cash_income_verif = "N" Then cash_income_verif = "N - No Verif Prvd"
'         If cash_income_verif = "?" Then cash_income_verif = "? - Delayed Verif"
'
'         If cash_expense_verif = "1" Then cash_expense_verif = "1 - Tax Returns"
'         If cash_expense_verif = "2" Then cash_expense_verif = "2 - Receipts"
'         If cash_expense_verif = "3" Then cash_expense_verif = "3 - Busi Records"
'         If cash_expense_verif = "6" Then cash_expense_verif = "6 - Other Doc"
'         If cash_expense_verif = "N" Then cash_expense_verif = "N - No Verif Prvd"
'         If cash_expense_verif = "?" Then cash_expense_verif = "? - Delayed Verif"
'
'
'         snap_retro_total_income = replace(snap_retro_total_income, "_", " ")
'         ' snap_retro_total_income = trim(snap_retro_total_income)
'
'         snap_retro_expenses = replace(snap_retro_expenses, "_", " ")
'         ' snap_retro_expenses = trim(snap_retro_expenses)
'
'         snap_prosp_total_income = replace(snap_prosp_total_income, "_", " ")
'         ' snap_prosp_total_income = trim(snap_prosp_total_income)
'
'         snap_prosp_expenses = replace(snap_prosp_expenses, "_", " ")
'         ' snap_prosp_expenses = trim(snap_prosp_expenses)
'
'         If snap_income_verif = "1" Then snap_income_verif = "1 - Tax Returns"
'         If snap_income_verif = "2" Then snap_income_verif = "2 - Receipts"
'         If snap_income_verif = "3" Then snap_income_verif = "3 - Busi Records"
'         If snap_income_verif = "4" Then snap_income_verif = "4 - Pend Out State"
'         If snap_income_verif = "6" Then snap_income_verif = "6 - Other Doc"
'         If snap_income_verif = "N" Then snap_income_verif = "N - No Verif Prvd"
'         If snap_income_verif = "?" Then snap_income_verif = "? - Delayed Verif"
'
'         If snap_expense_verif = "1" Then snap_expense_verif = "1 - Tax Returns"
'         If snap_expense_verif = "2" Then snap_expense_verif = "2 - Receipts"
'         If snap_expense_verif = "3" Then snap_expense_verif = "3 - Busi Records"
'         If snap_expense_verif = "4" Then snap_expense_verif = "4 - Pend Out State"
'         If snap_expense_verif = "6" Then snap_expense_verif = "6 - Other Doc"
'         If snap_expense_verif = "N" Then snap_expense_verif = "N - No Verif Prvd"
'         If snap_expense_verif = "?" Then snap_expense_verif = "? - Delayed Verif"
'
'         hc_method_a_total_income = replace(hc_method_a_total_income, "_", " ")
'         ' hc_method_a_total_income = trim(hc_method_a_total_income)
'
'         If hc_method_a_income_verif = "1" Then hc_method_a_income_verif = "1 - Tax Returns"
'         If hc_method_a_income_verif = "2" Then hc_method_a_income_verif = "2 - Receipts"
'         If hc_method_a_income_verif = "3" Then hc_method_a_income_verif = "3 - Busi Records"
'         If hc_method_a_income_verif = "6" Then hc_method_a_income_verif = "6 - Other Doc"
'         If hc_method_a_income_verif = "N" Then hc_method_a_income_verif = "N - No Verif Prvd"
'         If hc_method_a_income_verif = "?" Then hc_method_a_income_verif = "? - Delayed Verif"
'
'         hc_method_a_expenses = replace(hc_method_a_expenses, "_", " ")
'         ' hc_method_a_expenses = trim(hc_method_a_expenses)
'
'         If hc_method_a_expense_verif = "1" Then hc_method_a_expense_verif = "1 - Tax Returns"
'         If hc_method_a_expense_verif = "2" Then hc_method_a_expense_verif = "2 - Receipts"
'         If hc_method_a_expense_verif = "3" Then hc_method_a_expense_verif = "3 - Busi Records"
'         If hc_method_a_expense_verif = "6" Then hc_method_a_expense_verif = "6 - Other Doc"
'         If hc_method_a_expense_verif = "N" Then hc_method_a_expense_verif = "N - No Verif Prvd"
'         If hc_method_a_expense_verif = "?" Then hc_method_a_expense_verif = "? - Delayed Verif"
'
'         hc_method_b_total_income = replace(hc_method_b_total_income, "_", " ")
'         ' hc_method_b_total_income = trim(hc_method_b_total_income)
'
'         If hc_method_b_income_verif = "1" Then hc_method_b_income_verif = "1 - Tax Returns"
'         If hc_method_b_income_verif = "2" Then hc_method_b_income_verif = "2 - Receipts"
'         If hc_method_b_income_verif = "3" Then hc_method_b_income_verif = "3 - Busi Records"
'         If hc_method_b_income_verif = "6" Then hc_method_b_income_verif = "6 - Other Doc"
'         If hc_method_b_income_verif = "N" Then hc_method_b_income_verif = "N - No Verif Prvd"
'         If hc_method_b_income_verif = "?" Then hc_method_b_income_verif = "? - Delayed Verif"
'
'         hc_method_b_expenses = replace(hc_method_b_expenses, "_", " ")
'         ' hc_method_b_expenses = trim(hc_method_b_expenses)
'
'         If hc_method_b_expense_verif = "1" Then hc_method_b_expense_verif = "1 - Tax Returns"
'         If hc_method_b_expense_verif = "2" Then hc_method_b_expense_verif = "2 - Receipts"
'         If hc_method_b_expense_verif = "3" Then hc_method_b_expense_verif = "3 - Busi Records"
'         If hc_method_b_expense_verif = "6" Then hc_method_b_expense_verif = "6 - Other Doc"
'         If hc_method_b_expense_verif = "N" Then hc_method_b_expense_verif = "N - No Verif Prvd"
'         If hc_method_b_expense_verif = "?" Then hc_method_b_expense_verif = "? - Delayed Verif"
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_UNEA_panel(access_type, member_name, unea_type, unea_verif, panel_claim_nmbr, start_date, end_date, cola_amt, unea_amount, unea_pay_amount, unea_frequency, update_date, panel_ref_numb, pic_ave_inc, pic_prosp_income, retro_total)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen member_name, 2, 4, 33
'         EMReadScreen panel_type, 2, 5, 37
'         EMReadScreen panel_verif_code, 1, 5, 65
'         EMReadScreen panel_claim_nmbr, 15, 6, 37
'         EMReadScreen panel_start_date, 8, 7, 37
'         EMReadScreen panel_end_date, 8, 7, 68
'         EMReadScreen cola_disregard, 8, 10, 67
'         EMReadScreen update_date, 8, 21, 55
'
'         For unea_row = 17 to 13 Step -1
'             EMReadScreen pay_amount, 8, unea_row, 67
'             If pay_amount <> "________" Then
'                 unea_pay_amount = trim(pay_amount)
'                 Exit For
'             End If
'         Next
'         EMReadScreen total_amount, 8, 18, 68
'         EMReadScreen retro_total, 8, 18, 39
'
'         unea_amount = trim(total_amount)
'         retro_total = trim(retro_total)
'
'         EMWriteScreen "X", 10, 26       'opening SNAP pic
'         transmit
'         EMReadScreen pic_ave_inc, 8, 17, 56
'         EMReadScreen pic_prosp_income, 8, 18, 56
'         EMReadScreen panel_frequency_code, 1, 5, 64
'         PF3
'
'         If panel_type = "01" Then unea_type = "01 - RSDI, Disa"
'         If panel_type = "02" Then unea_type = "02 - RSDI, No Disa"
'         If panel_type = "03" Then unea_type = "03 - SSI"
'         If panel_type = "06" Then unea_type = "06 - Non-MN PA"
'         If panel_type = "11" Then unea_type = "11 - VA Disability"
'         If panel_type = "12" Then unea_type = "12 - VA Pension"
'         If panel_type = "13" Then unea_type = "13 - VA Other"
'         If panel_type = "38" Then unea_type = "38 - VA Aid & Attendance"
'         If panel_type = "14" Then unea_type = "14 - Unemployment Insurance"
'         If panel_type = "15" Then unea_type = "15 - Worker's Comp"
'         If panel_type = "16" Then unea_type = "16 - Railroad Retirement"
'         If panel_type = "17" Then unea_type = "17 - Other Retirement"
'         If panel_type = "18" Then unea_type = "18 - Military Enrirlement"
'         If panel_type = "19" Then unea_type = "19 - FC Child req FS"
'         If panel_type = "20" Then unea_type = "20 - FC Child not req FS"
'         If panel_type = "21" Then unea_type = "21 - FC Adult req FS"
'         If panel_type = "22" Then unea_type = "22 - FC Adult not req FS"
'         If panel_type = "23" Then unea_type = "23 - Dividends"
'         If panel_type = "24" Then unea_type = "24 - Interest"
'         If panel_type = "25" Then unea_type = "25 - Cnt gifts/prizes"
'         If panel_type = "26" Then unea_type = "26 - Strike Benefits"
'         If panel_type = "27" Then unea_type = "27 - Contract for Deed"
'         If panel_type = "28" Then unea_type = "28 - Illegal Income"
'         If panel_type = "29" Then unea_type = "29 - Other Countable"
'         If panel_type = "30" Then unea_type = "30 - Infrequent"
'         If panel_type = "31" Then unea_type = "31 - Other - FS Only"
'         If panel_type = "08" Then unea_type = "08 - Direct Child Support"
'         If panel_type = "35" Then unea_type = "35 - Direct Spousal Support"
'         If panel_type = "36" Then unea_type = "36 - Disbursed Child Support"
'         If panel_type = "37" Then unea_type = "37 - Disbursed Spousal Support"
'         If panel_type = "39" Then unea_type = "39 - Disbursed CS Arrears"
'         If panel_type = "40" Then unea_type = "40 - Disbursed Spsl Sup Arrears"
'         If panel_type = "43" Then unea_type = "43 - Disbursed Excess CS"
'         If panel_type = "44" Then unea_type = "44 - MSA - Excess Income for SSI"
'         If panel_type = "47" Then unea_type = "47 - Tribal Income"
'         If panel_type = "48" Then unea_type = "48 - Trust Income"
'         If panel_type = "49" Then unea_type = "49 - Non-Recurring"
'
'         If panel_verif_code = "1" Then unea_verif = "1 - Copy of Checks"
'         If panel_verif_code = "2" Then unea_verif = "2 - Award Letters"
'         If panel_verif_code = "3" Then unea_verif = "3 - System Initiated Verif"
'         If panel_verif_code = "4" Then unea_verif = "4 - Coltrl Stmt"
'         If panel_verif_code = "5" Then unea_verif = "5 - Pend Out State Verif"
'         If panel_verif_code = "6" Then unea_verif = "6 - Other Document"
'         If panel_verif_code = "7" Then unea_verif = "7 - Worker Initiated Verif"
'         If panel_verif_code = "8" Then unea_verif = "8 - RI Stubs"
'         If panel_verif_code = "N" Then unea_verif = "N - No Verif Prvd"
'         If panel_verif_code = "?" Then unea_verif = "? - Delayed Verif"
'
'         panel_claim_nmbr = replace(panel_claim_nmbr, "_", "")
'
'         start_date = replace(panel_start_date, " ", "/")
'         end_date = replace(panel_end_date, " ", "/")
'         If end_date = "__/__/__" Then end_date = ""
'         update_date = replace(update_date, " ", "/")
'         cola_amt = trim(cola_disregard)
'         If cola_amt = "________" Then cola_amt = ""
'
'         If panel_frequency_code = "1" Then unea_frequency = "1 - Monthly"
'         If panel_frequency_code = "2" Then unea_frequency = "2 - Semi Monthly"
'         If panel_frequency_code = "3" Then unea_frequency = "3 - Biweekly"
'         If panel_frequency_code = "4" Then unea_frequency = "4 - Weekly"
'         pic_ave_inc = trim(pic_ave_inc)
'         pic_prosp_income = trim(pic_prosp_income)
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_ACCT_panel(access_type, member_name, account_type, account_number, account_location, account_balance, account_verification, update_date, panel_ref_numb, balance_date, withdraw_penalty, withdraw_yn, withdraw_verif_code, count_cash, count_snap, count_hc, count_grh, count_ive, joint_own_yn, share_ratio, next_interest)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen member_name, 2, 4, 33
'         EMReadScreen panel_type, 2, 6, 44
'         EMReadScreen panel_number, 20, 7, 44
'         EMReadScreen panel_name, 20, 8, 44
'         EMReadScreen panel_balance, 8, 10, 46
'         EMReadScreen panel_verif_code, 1, 10, 64
'         EMReadScreen balance_date, 8, 11, 44
'         EMReadScreen withdraw_penalty, 8, 12, 46
'         EMReadScreen withdraw_yn, 1, 12, 64
'         EMReadScreen withdraw_verif_code, 1, 12, 72
'         EMReadScreen count_cash, 1, 14, 50
'         EMReadScreen count_snap, 1, 14, 57
'         EMReadScreen count_hc, 1, 14, 64
'         EMReadScreen count_grh, 1, 14, 72
'         EMReadScreen count_ive, 1, 14, 80
'         EMReadScreen joint_own_yn, 1, 15, 44
'         EMReadScreen share_ratio, 5, 15, 76
'         EMReadScreen next_interest, 5, 17, 57
'         EMReadScreen update_date, 8, 21, 55
'
'         If panel_type = "SV" Then account_type = "SV - Savings"
'         If panel_type = "CK" Then account_type = "CK - Checking"
'         If panel_type = "CE" Then account_type = "CE - Certificate of Deposit"
'         If panel_type = "MM" Then account_type = "MM - Money Market"
'         If panel_type = "DC" Then account_type = "DC - Debit Card"
'         If panel_type = "KO" Then account_type = "KO - Keogh Account"
'         If panel_type = "FT" Then account_type = "FT - Fed Thrift Savings Plan"
'         If panel_type = "SL" Then account_type = "SL - State & Local Govt"
'         If panel_type = "RA" Then account_type = "RA - Employee Ret Annuities"
'         If panel_type = "NP" Then account_type = "NP - Non-Profit Emmployee Ret"
'         If panel_type = "IR" Then account_type = "IR - Indiv Ret Acct"
'         If panel_type = "RH" Then account_type = "RH - Roth IRA"
'         If panel_type = "FR" Then account_type = "FR - Ret Plan for Employers"
'         If panel_type = "CT" Then account_type = "CT - Corp Ret Trust"
'         If panel_type = "RT" Then account_type = "RT - Other Ret Fund"
'         If panel_type = "QT" Then account_type = "QT - Qualified Tuition (529)"
'         If panel_type = "CA" Then account_type = "CA - Coverdell SV (530)"
'         If panel_type = "OE" Then account_type = "OE - Other Educational"
'         If panel_type = "OT" Then account_type = "OT - Other"
'
'         account_number = replace(panel_number, "_", "")
'         account_location =  replace(panel_name, "_", "")
'         account_balance = trim(panel_balance)
'
'         If panel_verif_code = "1"  Then account_verification = "1 - Bank Statement"
'         If panel_verif_code = "2"  Then account_verification = "2 - Agcy Ver Form"
'         If panel_verif_code = "3"  Then account_verification = "3 - Coltrl Contact"
'         If panel_verif_code = "5"  Then account_verification = "5 - Other Document"
'         If panel_verif_code = "6"  Then account_verification = "6 - Personal Statement"
'         If panel_verif_code = "N"  Then account_verification = "N - No Ver Prvd"
'
'         balance_date = replace(balance_date, " ", "/")
'         If balance_date = "__/__/__" Then balance_date = ""
'
'         withdraw_penalty = replace(withdraw_penalty, "_", "")
'         withdraw_penalty = trim(withdraw_penalty)
'         withdraw_yn = replace(withdraw_yn, "_", "")
'         If withdraw_verif_code = "1"  Then withdraw_verif_code = "1 - Bank Statement"
'         If withdraw_verif_code = "2"  Then withdraw_verif_code = "2 - Agcy Ver Form"
'         If withdraw_verif_code = "3"  Then withdraw_verif_code = "3 - Coltrl Contact"
'         If withdraw_verif_code = "5"  Then withdraw_verif_code = "5 - Other Document"
'         If withdraw_verif_code = "6"  Then withdraw_verif_code = "6 - Personal Statement"
'         If withdraw_verif_code = "N"  Then withdraw_verif_code = "N - No Ver Prvd"
'
'         count_cash = replace(count_cash, "_", "")
'         count_snap = replace(count_snap, "_", "")
'         count_hc = replace(count_hc, "_", "")
'         count_grh = replace(count_grh, "_", "")
'         count_ive = replace(count_ive, "_", "")
'
'         share_ratio = replace(share_ratio, " ", "")
'
'         next_interest = replace(next_interest, " ", "/")
'         If next_interest = "__/__" Then next_interest = ""
'
'         update_date = replace(update_date, " ", "/")
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_CARS_panel(access_type, member_name, cars_type, cars_year, cars_make, cars_model, cars_verif, update_date, panel_ref_numb, cars_trade_in, cars_loan, cars_source, cars_owed, cars_owed_verif_code, cars_owed_date, cars_use, cars_hc_benefit, cars_joint_yn, cars_share)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen member_name, 2, 4, 33
'         EMReadScreen cars_type, 1, 6, 43
'         EMReadScreen cars_year, 4, 8, 31
'         EMReadScreen cars_make, 15, 8, 43
'         EMReadScreen cars_model, 15, 8, 66
'         EMReadScreen cars_trade_in, 8, 9, 45            'not output
'         EMReadScreen cars_loan, 8, 9, 62                'not output
'         EMReadScreen cars_source, 1, 9, 80              'not output
'         EMReadScreen cars_verif_code, 1, 10, 60
'         EMReadScreen cars_owed, 8, 12, 45               'not output
'         EMReadScreen cars_owed_verif_code, 1, 12, 60    'not output
'         EMReadScreen cars_owed_date, 8, 13, 43          'not output
'         EMReadScreen cars_use, 1, 15, 43                'not output
'         EMReadScreen cars_hc_benefit, 1, 15, 76         'not output
'         EMReadScreen cars_joint_yn, 1, 16, 43           'not output
'         EMReadScreen cars_share, 5, 16, 76              'not output
'         EMReadScreen cars_update, 8, 21, 55
'
'         If cars_type = "1" Then cars_type = "1 - Car"
'         If cars_type = "2" Then cars_type = "2 - Truck"
'         If cars_type = "3" Then cars_type = "3 - Van"
'         If cars_type = "4" Then cars_type = "4 - Camper"
'         If cars_type = "5" Then cars_type = "5 - Motorcycle"
'         If cars_type = "6" Then cars_type = "6 - Trailer"
'         If cars_type = "7" Then cars_type = "7 - Other"
'
'         cars_make = replace(cars_make, "_", "")
'         cars_model = replace(cars_model, "_", "")
'
'
'         cars_trade_in = replace(cars_trade_in, "_", "")
'         cars_trade_in = trim(cars_trade_in)
'
'         cars_loan = replace(cars_loan, "_", "")
'         cars_loan = trim(cars_loan)
'
'         If cars_source = "1" Then cars_source = "1 - NADA"
'         If cars_source = "2" Then cars_source = "2 - Appraisal Val"
'         If cars_source = "3" Then cars_source = "3 - Client Stmt"
'         If cars_source = "4" Then cars_source = "4 - Other Document"
'
'         If cars_verif_code = "1" Then cars_verif = "1 - Title"
'         If cars_verif_code = "2" Then cars_verif = "2 - License Reg"
'         If cars_verif_code = "3" Then cars_verif = "3 - DMV"
'         If cars_verif_code = "4" Then cars_verif = "4 - Purchase Agmt"
'         If cars_verif_code = "5" Then cars_verif = "5 - Other Document"
'         If cars_verif_code = "N" Then cars_verif = "N - No Ver Prvd"
'
'         cars_owed = replace(cars_owed, "_", "")
'         cars_owed = trim(cars_owed)
'
'         If cars_owed_verif_code = "1" Then cars_owed_verif_code = "1 - Bank/Lending Inst Stmt"
'         If cars_owed_verif_code = "2" Then cars_owed_verif_code = "2 - Private Lender Stmt"
'         If cars_owed_verif_code = "3" Then cars_owed_verif_code = "3 - Other Document"
'         If cars_owed_verif_code = "4" Then cars_owed_verif_code = "4 - Pend Out State Verif"
'         If cars_owed_verif_code = "N" Then cars_owed_verif_code = "N - No Ver Prvd"
'
'         cars_owed_date = replace(cars_owed_date, " ", "/")
'         If cars_owed_date = "__/__/__" Then cars_owed_date = ""
'
'         If cars_use = "1" Then cars_use = "1 - Primary Vehicle"
'         If cars_use = "2" Then cars_use = "2 - Employment/Training Search"
'         If cars_use = "3" Then cars_use = "3 - Disa Transportation"
'         If cars_use = "4" Then cars_use = "4 - Income Producing"
'         If cars_use = "5" Then cars_use = "5 - Used as Home"
'         If cars_use = "7" Then cars_use = "7 - Unlicensed"
'         If cars_use = "8" Then cars_use = "8 - Other Countable"
'         If cars_use = "9" Then cars_use = "9 - Unavailable"
'         If cars_use = "0" Then cars_use = "0 - Long Distance Employment Travel"
'         If cars_use = "A" Then cars_use = "A - Carry Heating Fuel or Water"
'
'         cars_hc_benefit = replace(cars_hc_benefit, "_", "")
'         cars_joint_yn = replace(cars_joint_yn, "_", "")
'         cars_share = replace(cars_share, " ", "")
'
'         update_date = replace(cars_update, " ", "/")
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_FACI_panel(access_type, notes_on_faci, facility_name, facility_vendor_number, facility_type, facility_FS_elig, FS_facility_type, facility_waiver_type, facility_LTC_inelig_reason, facility_inelig_begin_date, facility_inelig_end_date, facility_anticipated_out_date, facility_GRH_plan_required, facility_GRH_plan_verif, facility_cty_app_place, facility_approval_cty_name, facility_GRH_DOC_amount, facility_GRH_postpay, facility_stay_one_rate, facility_stay_one_date_in, facility_stay_one_date_out, facility_stay_two_rate, facility_stay_two_date_in, facility_stay_two_date_out, facility_stay_three_rate, facility_stay_three_date_in, facility_stay_three_date_out, facility_stay_four_rate, facility_stay_four_date_in, facility_stay_four_date_out, facility_stay_five_rate, facility_stay_five_date_in, facility_stay_five_date_out)
'     If access_type = "READ" Then
'         EMReadScreen facility_name,                 30, 6, 43
'         EMReadScreen facility_vendor_number,        8, 5, 43
'         EMReadScreen facility_type,                 2, 7, 43
'         EMReadScreen facility_FS_elig,              1, 8, 43
'         EMReadScreen FS_facility_type,              1, 8, 71
'         EMReadScreen facility_waiver_type,          2, 7, 71
'         EMReadScreen facility_LTC_inelig_reason,    1,  9, 43
'         EMReadScreen facility_inelig_begin_date,    10, 10, 52
'         EMReadScreen facility_inelig_end_date,      10, 10, 71
'         EMReadScreen facility_anticipated_out_date, 10, 9, 71
'
'         facility_name = replace(facility_name, "_", "")
'         If facility_type = "41" Then facility_type = "41 - NF-I"
'         If facility_type = "42" Then facility_type = "42 - NF-II"
'         If facility_type = "43" Then facility_type = "43 - ICF-DD"
'         If facility_type = "44" Then facility_type = "44 - Short Stay In NF-I"
'         If facility_type = "45" Then facility_type = "45 - Short Stay In NF-II"
'         If facility_type = "46" Then facility_type = "46 - Short Stay in ICF-DD"
'         If facility_type = "47" Then facility_type = "47 - RTC - Not IMD"
'         If facility_type = "48" Then facility_type = "48 - Medical Hospital"
'         If facility_type = "49" Then facility_type = "49 - MSOP"
'         If facility_type = "50" Then facility_type = "50 - IMD/RTC"
'         If facility_type = "51" Then facility_type = "51 - Rule 31 CD-IMD"
'         If facility_type = "52" Then facility_type = "52 - Rule 36 MI-IMD"
'         If facility_type = "53" Then facility_type = "53 - IMD Hospitals"
'         If facility_type = "55" Then facility_type = "55 - Adult Foster Care/Rule 203"
'         If facility_type = "56" Then facility_type = "56 - GRH (Not FC or Rule 36)"
'         If facility_type = "57" Then facility_type = "57 - Rule 36 MI-Non-IMD"
'         If facility_type = "60" Then facility_type = "60 - Non-GRH"
'         If facility_type = "61" Then facility_type = "61 - Rule 31 CD-Non-IMD"
'         If facility_type = "67" Then facility_type = "67 - Family Violence Shelter"
'         If facility_type = "68" Then facility_type = "68 - County Correctional Facility"
'         If facility_type = "69" Then facility_type = "69 - Non-Cty Adult Correctional"
'
'         If FS_facility_type = "1" Then FS_facility_type = "1 - Fed Subsidized Housing for Elderly"
'         If FS_facility_type = "2" Then FS_facility_type = "2 - Licensed Facility/Treatment Center - CD"
'         If FS_facility_type = "3" Then FS_facility_type = "3 - Blind or Disabled RSDI/SSI Recipient"
'         If FS_facility_type = "4" Then FS_facility_type = "4 - Family Violence Shelter"
'         If FS_facility_type = "5" Then FS_facility_type = "5 - Temporary Shelter for Homeless"
'         If FS_facility_type = "6" Then FS_facility_type = "6 - Not a facility by FS Definition"
'
'         If facility_waiver_type = "01" Then facility_waiver_type = "01 - CADI"
'         If facility_waiver_type = "02" Then facility_waiver_type = "02 - CAC"
'         If facility_waiver_type = "03" Then facility_waiver_type = "03 - EW Single"
'         If facility_waiver_type = "04" Then facility_waiver_type = "04 - EW Married"
'         If facility_waiver_type = "05" Then facility_waiver_type = "05 - TBI"
'         If facility_waiver_type = "06" Then facility_waiver_type = "06 - DD"
'         If facility_waiver_type = "07" Then facility_waiver_type = "07 - ACS (Alt Care Services DD)"
'         If facility_waiver_type = "08" Then facility_waiver_type = "08 - SISEW Single"
'         If facility_waiver_type = "09" Then facility_waiver_type = "09 - SISEW Married"
'
'         If facility_LTC_inelig_reason = "L" Then facility_LTC_inelig_reason = "L - This level of Care Not Required"
'         If facility_LTC_inelig_reason = "N" Then facility_LTC_inelig_reason = "N - Not Pre-Screened"
'         If facility_LTC_inelig_reason = "_" Then facility_LTC_inelig_reason = ""
'
'         facility_inelig_begin_date = replace(facility_inelig_begin_date, " ", "/")
'         If facility_inelig_begin_date = "__/__/____" Then facility_inelig_begin_date = ""
'         facility_inelig_end_date = replace(facility_inelig_end_date, " ", "/")
'         If facility_inelig_end_date = "__/__/____" Then facility_inelig_end_date = ""
'         facility_anticipated_out_date = replace(facility_anticipated_out_date, " ", "/")
'         If facility_anticipated_out_date = "__/__/____" Then facility_anticipated_out_date = ""
'
'         EMReadScreen facility_GRH_plan_required,    1, 11, 52
'         EMReadScreen facility_cty_app_place,        1, 12, 52
'         EMReadScreen facility_GRH_plan_verif,       1, 11, 71
'         EMReadScreen facility_approval_cty,         2, 12, 71
'         EMReadScreen facility_GRH_DOC_amount,       8, 13, 45
'         EMReadScreen facility_GRH_postpay,          1, 13, 71
'
'         EMReadScreen facility_stay_one_rate,        1,  14, 34
'         EMReadScreen facility_stay_one_date_in,     10, 14, 47
'         EMReadScreen facility_stay_one_date_out,    10, 14, 71
'
'         EMReadScreen facility_stay_two_rate,        1,  15, 34
'         EMReadScreen facility_stay_two_date_in,     10, 15, 47
'         EMReadScreen facility_stay_two_date_out,    10, 15, 71
'
'         EMReadScreen facility_stay_three_rate,      1,  16, 34
'         EMReadScreen facility_stay_three_date_in,   10, 16, 47
'         EMReadScreen facility_stay_three_date_out,  10, 16, 71
'
'         EMReadScreen facility_stay_four_rate,       1,  17, 34
'         EMReadScreen facility_stay_four_date_in,    10, 17, 47
'         EMReadScreen facility_stay_four_date_out,   10, 17, 71
'
'         EMReadScreen facility_stay_five_rate,       1,  18, 34
'         EMReadScreen facility_stay_five_date_in,    10, 18, 47
'         EMReadScreen facility_stay_five_date_out,   10, 18, 71
'
'         facility_GRH_plan_required = replace(facility_GRH_plan_required, "_", "")
'         facility_GRH_plan_verif = replace(facility_GRH_plan_verif, "_", "")
'         facility_cty_app_place = replace(facility_cty_app_place, "_", "")
'         Call get_county_name_from_county_code(facility_approval_cty, facility_approval_cty_name, TRUE)
'         facility_GRH_DOC_amount = replace(facility_GRH_DOC_amount, "_", "")
'         facility_GRH_DOC_amount = trim(facility_GRH_DOC_amount)
'         facility_GRH_postpay = replace(facility_GRH_postpay, "_", "")
'
'         If facility_stay_one_rate = "1" Then facility_stay_one_rate = "Rate 1"
'         If facility_stay_one_rate = "2" Then facility_stay_one_rate = "Rate 2"
'         If facility_stay_one_rate = "3" Then facility_stay_one_rate = "Rate 3"
'         If facility_stay_one_rate = "_" Then facility_stay_one_rate = "      "
'         facility_stay_one_date_in = replace(facility_stay_one_date_in, " ", "/")
'         If facility_stay_one_date_in = "__/__/____" Then facility_stay_one_date_in = ""
'         facility_stay_one_date_out = replace(facility_stay_one_date_out, " ", "/")
'         If facility_stay_one_date_out = "__/__/____" Then facility_stay_one_date_out = ""
'
'         If facility_stay_two_rate = "1" Then facility_stay_two_rate = "Rate 1"
'         If facility_stay_two_rate = "2" Then facility_stay_two_rate = "Rate 2"
'         If facility_stay_two_rate = "3" Then facility_stay_two_rate = "Rate 3"
'         If facility_stay_two_rate = "_" Then facility_stay_two_rate = "      "
'         facility_stay_two_date_in = replace(facility_stay_two_date_in, " ", "/")
'         If facility_stay_two_date_in = "__/__/____" Then facility_stay_two_date_in = ""
'         facility_stay_two_date_out = replace(facility_stay_two_date_out, " ", "/")
'         If facility_stay_two_date_out = "__/__/____" Then facility_stay_two_date_out = ""
'
'         If facility_stay_three_rate = "1" Then facility_stay_three_rate = "Rate 1"
'         If facility_stay_three_rate = "2" Then facility_stay_three_rate = "Rate 2"
'         If facility_stay_three_rate = "3" Then facility_stay_three_rate = "Rate 3"
'         If facility_stay_three_rate = "_" Then facility_stay_three_rate = "      "
'         facility_stay_three_date_in = replace(facility_stay_three_date_in, " ", "/")
'         If facility_stay_three_date_in = "__/__/____" Then facility_stay_three_date_in = ""
'         facility_stay_three_date_out = replace(facility_stay_three_date_out, " ", "/")
'         If facility_stay_three_date_out = "__/__/____" Then facility_stay_three_date_out = ""
'
'         If facility_stay_four_rate = "1" Then facility_stay_four_rate = "Rate 1"
'         If facility_stay_four_rate = "2" Then facility_stay_four_rate = "Rate 2"
'         If facility_stay_four_rate = "3" Then facility_stay_four_rate = "Rate 3"
'         If facility_stay_four_rate = "_" Then facility_stay_four_rate = "      "
'         facility_stay_four_date_in = replace(facility_stay_four_date_in, " ", "/")
'         If facility_stay_four_date_in = "__/__/____" Then facility_stay_four_date_in = ""
'         facility_stay_four_date_out = replace(facility_stay_four_date_out, " ", "/")
'         If facility_stay_four_date_out = "__/__/____" Then facility_stay_four_date_out = ""
'
'         If facility_stay_five_rate = "1" Then facility_stay_five_rate = "Rate 1"
'         If facility_stay_five_rate = "2" Then facility_stay_five_rate = "Rate 2"
'         If facility_stay_five_rate = "3" Then facility_stay_five_rate = "Rate 3"
'         If facility_stay_five_rate = "_" Then facility_stay_five_rate = "      "
'         facility_stay_five_date_in = replace(facility_stay_five_date_in, " ", "/")
'         If facility_stay_five_date_in = "__/__/____" Then facility_stay_five_date_in = ""
'         facility_stay_five_date_out = replace(facility_stay_five_date_out, " ", "/")
'         If facility_stay_five_date_out = "__/__/____" Then facility_stay_five_date_out = ""
'     End If
' end function
'
' function access_SECU_panel(access_type, member_name, security_type, security_account_number, security_name, security_cash_value, security_verif, secu_update_date, panel_ref_numb, security_face_value, security_withdraw, security_withdraw_yn, security_withdraw_verif, secu_cash_yn, secu_snap_yn, secu_hc_yn, secu_grh_yn, secu_ive_yn, secu_joint, secu_ratio, security_eff_date)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen member_name, 2, 4, 33
'         EMReadScreen panel_type, 2, 6, 50
'         EMReadScreen security_account_number, 12, 7, 50
'         EMReadScreen security_name, 20, 8, 50
'         EMReadScreen security_cash_value, 8, 10, 52
'         EMReadScreen security_eff_date, 8, 11, 35   'not output
'         EMReadScreen verif_code, 1, 11, 50
'         EMReadScreen security_face_value, 8, 12, 52     'not output
'         EMReadScreen security_withdraw, 8, 13, 52       'not output
'         EMReadScreen security_withdraw_yn, 1, 13, 72    'not output
'         EMReadScreen security_withdraw_verif, 1, 13, 80 'not output
'
'         EMReadScreen secu_cash_yn, 1, 15, 50    'not output
'         EMReadScreen secu_snap_yn, 1, 15, 57    'not output
'         EMReadScreen secu_hc_yn, 1, 15, 64      'not output
'         EMReadScreen secu_grh_yn, 1, 15, 72     'not output
'         EMReadScreen secu_ive_yn, 1, 15, 80     'not output
'
'         EMReadScreen secu_joint, 1, 16, 44      'not output
'         EMReadScreen secu_ratio, 5, 16, 76      'not output
'         EMReadScreen secu_update_date, 8, 21, 55
'
'         If panel_type = "LI" Then security_type = "LI - Life Insurance"
'         If panel_type = "ST" Then security_type = "ST - Stocks"
'         If panel_type = "BO" Then security_type = "BO - Bonds"
'         If panel_type = "CD" Then security_type = "CD - Ctrct for Deed"
'         If panel_type = "MO" Then security_type = "MO - Mortgage Note"
'         If panel_type = "AN" Then security_type = "AN - Annuity"
'         If panel_type = "OT" Then security_type = "OT - Other"
'
'         security_account_number = replace(security_account_number, "_", "")
'         security_name = replace(security_name, "_", "")
'
'         security_cash_value = replace(security_cash_value, "_", "")
'         security_cash_value = trim(security_cash_value)
'
'         security_eff_date = replace(security_eff_date, " ", "/")
'         If security_eff_date = "__/__/__" Then security_eff_date = ""
'
'         If verif_code = "1" Then security_verif = "1 - Agency Form"
'         If verif_code = "2" Then security_verif = "2 - Source Doc"
'         If verif_code = "3" Then security_verif = "3 - Phone Contact"
'         If verif_code = "5" Then security_verif = "5 - Other Document"
'         If verif_code = "6" Then security_verif = "6 - Personal Statement"
'         If verif_code = "N" Then security_verif = "N - No Ver Prov"
'
'         security_face_value = replace(security_face_value, "_", "")
'         security_face_value = trim(security_face_value)
'
'         security_withdraw = replace(security_withdraw, "_", "")
'         security_withdraw = trim(security_withdraw)
'
'         security_withdraw_yn = replace(security_withdraw_yn, "_", "")
'
'         If security_withdraw_verif = "1" Then security_withdraw_verif = "1 - Agency Form"
'         If security_withdraw_verif = "2" Then security_withdraw_verif = "2 - Source Doc"
'         If security_withdraw_verif = "3" Then security_withdraw_verif = "3 - Phone Contact"
'         If security_withdraw_verif = "4" Then security_withdraw_verif = "4 - Other Document"
'         If security_withdraw_verif = "5" Then security_withdraw_verif = "5 - Personal Stmt"
'         If security_withdraw_verif = "N" Then security_withdraw_verif = "N - No Ver Prov"
'
'         secu_cash_yn = replace(secu_cash_yn, "_", "")
'         secu_snap_yn = replace(secu_snap_yn, "_", "")
'         secu_hc_yn = replace(secu_hc_yn, "_", "")
'         secu_grh_yn = replace(secu_grh_yn, "_", "")
'         secu_ive_yn = replace(secu_ive_yn, "_", "")
'
'         secu_joint = replace(secu_joint, "_", "")
'         secu_ratio = replace(secu_ratio, " ", "")
'
'         secu_update_date = replace(secu_update_date, " ", "/")
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_SHEL_panel(access_type, hud_sub_yn, shared_yn, paid_to, rent_retro_amt, rent_retro_verif, rent_prosp_amt, rent_prosp_verif, lot_rent_retro_amt, lot_rent_retro_verif, lot_rent_prosp_amt, lot_rent_prosp_verif, mortgage_retro_amt, mortgage_retro_verif, mortgage_prosp_amt, mortgage_prosp_verif, insurance_retro_amt, insurance_retro_verif, insurance_prosp_amt, insurance_prosp_verif, tax_retro_amt, tax_retro_verif, tax_prosp_amt, tax_prosp_verif, room_retro_amt, room_retro_verif, room_prosp_amt, room_prosp_verif, garage_retro_amt, garage_retro_verif, garage_prosp_amt, garage_prosp_verif, subsidy_retro_amt, subsidy_retro_verif, subsidy_prosp_amt, subsidy_prosp_verif)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen hud_sub_yn,            1, 6, 46
'         EMReadScreen shared_yn,             1, 6, 64
'         EMReadScreen paid_to,               25, 7, 50
'
'         paid_to = replace(paid_to, "_", "")
'
'         EMReadScreen rent_retro_amt,        8, 11, 37
'         EMReadScreen rent_retro_verif,      2, 11, 48
'         EMReadScreen rent_prosp_amt,        8, 11, 56
'         EMReadScreen rent_prosp_verif,      2, 11, 67
'
'         rent_retro_amt = replace(rent_retro_amt, "_", "")
'         rent_retro_amt = trim(rent_retro_amt)
'         If rent_retro_verif = "SF" Then rent_retro_verif = "SF - Shelter Form"
'         If rent_retro_verif = "LE" Then rent_retro_verif = "LE - Lease"
'         If rent_retro_verif = "RE" Then rent_retro_verif = "RE - Rent Receipt"
'         If rent_retro_verif = "OT" Then rent_retro_verif = "OT - Other Document"
'         If rent_retro_verif = "NC" Then rent_retro_verif = "NC - Chg Rept, Neg Impact"
'         If rent_retro_verif = "PC" Then rent_retro_verif = "PC - Chg Rept, Pos Imact"
'         If rent_retro_verif = "NO" Then rent_retro_verif = "NO - No Ver Prvd"
'         If rent_retro_verif = "__" Then rent_retro_verif = ""
'         rent_prosp_amt = replace(rent_prosp_amt, "_", "")
'         rent_prosp_amt = trim(rent_prosp_amt)
'         If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
'         If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
'         If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
'         If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Document"
'         If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Ver Prvd"
'         If rent_prosp_verif = "__" Then rent_prosp_verif = ""
'
'         EMReadScreen lot_rent_retro_amt,    8, 12, 37
'         EMReadScreen lot_rent_retro_verif,  2, 12, 48
'         EMReadScreen lot_rent_prosp_amt,    8, 12, 56
'         EMReadScreen lot_rent_prosp_verif,  2, 12, 67
'
'         lot_rent_retro_amt = replace(lot_rent_retro_amt, "_", "")
'         lot_rent_retro_amt = trim(lot_rent_retro_amt)
'         If lot_rent_retro_verif = "LE" Then lot_rent_retro_verif = "LE - Lease"
'         If lot_rent_retro_verif = "RE" Then lot_rent_retro_verif = "RE - Rent Receipt"
'         If lot_rent_retro_verif = "BI" Then lot_rent_retro_verif = "BI - Billing Stmt"
'         If lot_rent_retro_verif = "OT" Then lot_rent_retro_verif = "OT - Other Document"
'         If lot_rent_retro_verif = "NC" Then lot_rent_retro_verif = "NC - Chg Rept, Neg Impact"
'         If lot_rent_retro_verif = "PC" Then lot_rent_retro_verif = "PC - Chg Rept, Pos Imact"
'         If lot_rent_retro_verif = "NO" Then lot_rent_retro_verif = "NO - No Ver Prvd"
'         If lot_rent_retro_verif = "__" Then lot_rent_retro_verif = ""
'         lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
'         lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
'         If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
'         If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
'         If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
'         If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Document"
'         If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Ver Prvd"
'         If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""
'
'         EMReadScreen mortgage_retro_amt,    8, 13, 37
'         EMReadScreen mortgage_retro_verif,  2, 13, 48
'         EMReadScreen mortgage_prosp_amt,    8, 13, 56
'         EMReadScreen mortgage_prosp_verif,  2, 13, 67
'
'         mortgage_retro_amt = replace(mortgage_retro_amt, "_", "")
'         mortgage_retro_amt = trim(mortgage_retro_amt)
'         If mortgage_retro_verif = "MO" Then mortgage_retro_verif = "MO - Mortgage Pmt Book"
'         If mortgage_retro_verif = "CD" Then mortgage_retro_verif = "CD - Ctrct fro Deed"
'         If mortgage_retro_verif = "OT" Then mortgage_retro_verif = "OT - Other Document"
'         If mortgage_retro_verif = "NC" Then mortgage_retro_verif = "NC - Chg Rept, Neg Impact"
'         If mortgage_retro_verif = "PC" Then mortgage_retro_verif = "PC - Chg Rept, Pos Imact"
'         If mortgage_retro_verif = "NO" Then mortgage_retro_verif = "NO - No Ver Prvd"
'         If mortgage_retro_verif = "__" Then mortgage_retro_verif = ""
'         mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
'         mortgage_prosp_amt = trim(mortgage_prosp_amt)
'         If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
'         If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
'         If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Document"
'         If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Ver Prvd"
'         If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""
'
'         EMReadScreen insurance_retro_amt,   8, 14, 37
'         EMReadScreen insurance_retro_verif, 2, 14, 48
'         EMReadScreen insurance_prosp_amt,   8, 14, 56
'         EMReadScreen insurance_prosp_verif, 2, 14, 67
'
'         insurance_retro_amt = replace(insurance_retro_amt, "_", "")
'         insurance_retro_amt = trim(insurance_retro_amt)
'         If insurance_retro_verif = "BI" Then insurance_retro_verif = "BI - Billing Stmt"
'         If insurance_retro_verif = "OT" Then insurance_retro_verif = "OT - Other Document"
'         If insurance_retro_verif = "NC" Then insurance_retro_verif = "NC - Chg Rept, Neg Impact"
'         If insurance_retro_verif = "PC" Then insurance_retro_verif = "PC - Chg Rept, Pos Imact"
'         If insurance_retro_verif = "NO" Then insurance_retro_verif = "NO - No Ver Prvd"
'         If insurance_retro_verif = "__" Then insurance_retro_verif = ""
'         insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
'         insurance_prosp_amt = trim(insurance_prosp_amt)
'         If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
'         If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Document"
'         If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Ver Prvd"
'         If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""
'
'         EMReadScreen tax_retro_amt,         8, 15, 37
'         EMReadScreen tax_retro_verif,       2, 15, 48
'         EMReadScreen tax_prosp_amt,         8, 15, 56
'         EMReadScreen tax_prosp_verif,       2, 15, 67
'
'         tax_retro_amt = replace(tax_retro_amt, "_", "")
'         tax_retro_amt = trim(tax_retro_amt)
'         If tax_retro_verif = "TX" Then tax_retro_verif = "TX - Prop Tax Stmt"
'         If tax_retro_verif = "OT" Then tax_retro_verif = "OT - Other Document"
'         If tax_retro_verif = "NC" Then tax_retro_verif = "NC - Chg Rept, Neg Impact"
'         If tax_retro_verif = "PC" Then tax_retro_verif = "PC - Chg Rept, Pos Imact"
'         If tax_retro_verif = "NO" Then tax_retro_verif = "NO - No Ver Prvd"
'         If tax_retro_verif = "__" Then tax_retro_verif = ""
'         tax_prosp_amt = replace(tax_prosp_amt, "_", "")
'         tax_prosp_amt = trim(tax_prosp_amt)
'         If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
'         If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Document"
'         If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Ver Prvd"
'         If tax_prosp_verif = "__" Then tax_prosp_verif = ""
'
'         EMReadScreen room_retro_amt,        8, 16, 37
'         EMReadScreen room_retro_verif,      2, 16, 48
'         EMReadScreen room_prosp_amt,        8, 16, 56
'         EMReadScreen room_prosp_verif,      2, 16, 67
'
'         room_retro_amt = replace(room_retro_amt, "_", "")
'         room_retro_amt = trim(room_retro_amt)
'         If room_retro_verif = "SF" Then room_retro_verif = "SF - Shelter Form"
'         If room_retro_verif = "LE" Then room_retro_verif = "LE - Lease"
'         If room_retro_verif = "RE" Then room_retro_verif = "RE - Rent Receipt"
'         If room_retro_verif = "OT" Then room_retro_verif = "OT - Other Document"
'         If room_retro_verif = "NC" Then room_retro_verif = "NC - Chg Rept, Neg Impact"
'         If room_retro_verif = "PC" Then room_retro_verif = "PC - Chg Rept, Pos Imact"
'         If room_retro_verif = "NO" Then room_retro_verif = "NO - No Ver Prvd"
'         If room_retro_verif = "__" Then room_retro_verif = ""
'         room_prosp_amt = replace(room_prosp_amt, "_", "")
'         room_prosp_amt = trim(room_prosp_amt)
'         If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
'         If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
'         If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
'         If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Document"
'         If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Ver Prvd"
'         If room_prosp_verif = "__" Then room_prosp_verif = ""
'
'         EMReadScreen garage_retro_amt,      8, 17, 37
'         EMReadScreen garage_retro_verif,    2, 17, 48
'         EMReadScreen garage_prosp_amt,      8, 17, 56
'         EMReadScreen garage_prosp_verif,    2, 17, 67
'
'         garage_retro_amt = replace(garage_retro_amt, "_", "")
'         garage_retro_amt = trim(garage_retro_amt)
'         If garage_retro_verif = "SF" Then garage_retro_verif = "SF - Shelter Form"
'         If garage_retro_verif = "LE" Then garage_retro_verif = "LE - Lease"
'         If garage_retro_verif = "RE" Then garage_retro_verif = "RE - Rent Receipt"
'         If garage_retro_verif = "OT" Then garage_retro_verif = "OT - Other Document"
'         If garage_retro_verif = "NC" Then garage_retro_verif = "NC - Chg Rept, Neg Impact"
'         If garage_retro_verif = "PC" Then garage_retro_verif = "PC - Chg Rept, Pos Imact"
'         If garage_retro_verif = "NO" Then garage_retro_verif = "NO - No Ver Prvd"
'         If garage_retro_verif = "__" Then garage_retro_verif = ""
'         garage_prosp_amt = replace(garage_prosp_amt, "_", "")
'         garage_prosp_amt = trim(garage_prosp_amt)
'         If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
'         If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
'         If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
'         If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Document"
'         If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg Rept, Neg Impact"
'         If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg Rept, Pos Imact"
'         If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Ver Prvd"
'         If garage_prosp_verif = "__" Then garage_prosp_verif = ""
'
'         EMReadScreen subsidy_retro_amt,     8, 18, 37
'         EMReadScreen subsidy_retro_verif,   2, 18, 48
'         EMReadScreen subsidy_prosp_amt,     8, 18, 56
'         EMReadScreen subsidy_prosp_verif,   2, 18, 67
'
'         subsidy_retro_amt = replace(subsidy_retro_amt, "_", "")
'         subsidy_retro_amt = trim(subsidy_retro_amt)
'         If subsidy_retro_verif = "SF" Then subsidy_retro_verif = "SF - Shelter Form"
'         If subsidy_retro_verif = "LE" Then subsidy_retro_verif = "LE - Lease"
'         If subsidy_retro_verif = "OT" Then subsidy_retro_verif = "OT - Other Document"
'         If subsidy_retro_verif = "NO" Then subsidy_retro_verif = "NO - No Ver Prvd"
'         If subsidy_retro_verif = "__" Then subsidy_retro_verif = ""
'         subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
'         subsidy_prosp_amt = trim(subsidy_prosp_amt)
'         If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
'         If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
'         If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Document"
'         If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Ver Prvd"
'         If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""
'     End If
' end function
'
' function access_REST_panel(access_type, member_name, rest_type, rest_verif, rest_update_date, panel_ref_numb, rest_market_value, value_verif_code, rest_amt_owed, amt_owed_verif_code, rest_eff_date, rest_status, rest_joint_yn, rest_ratio, repymt_agree_date)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen member_name, 2, 4, 33
'         EMReadScreen type_code, 1, 6, 39
'         EMReadScreen type_verif_code, 2, 6, 62
'         EMReadScreen rest_market_value, 10, 8, 41
'         EMReadScreen value_verif_code, 2, 8, 62
'         EMReadScreen rest_amt_owed, 10, 9, 41
'         EMReadScreen amt_owed_verif_code, 2, 9, 62
'         EMReadScreen rest_eff_date, 8, 10, 39
'         EMReadScreen rest_status, 1, 12, 54
'         EMReadScreen rest_joint_yn, 1, 13, 54
'         EMReadScreen rest_ratio, 5, 14, 54
'         EMReadScreen repymt_agree_date, 8, 16, 62
'         EMReadScreen rest_update_date, 8, 21, 55
'
'         If type_code = "1" Then rest_type = "1 - House"
'         If type_code = "2" Then rest_type = "2 - Land"
'         If type_code = "3" Then rest_type = "3 - Buildings"
'         If type_code = "4" Then rest_type = "4 - Mobile Home"
'         If type_code = "5" Then rest_type = "5 - Life Estate"
'         If type_code = "6" Then rest_type = "6 - Other"
'
'         If type_verif_code = "TX" Then rest_verif = "TX - Property Tax Statement"
'         If type_verif_code = "PU" Then rest_verif = "PU - Purchase Agreement"
'         If type_verif_code = "TI" Then rest_verif = "TI - Title/Deed"
'         If type_verif_code = "CD" Then rest_verif = "CD - Contract for Deed"
'         If type_verif_code = "CO" Then rest_verif = "CO - County Record"
'         If type_verif_code = "OT" Then rest_verif = "OT - Other Document"
'         If type_verif_code = "NO" Then rest_verif = "NO - No Ver Prvd"
'
'         rest_market_value = replace(rest_market_value, "_", "")
'         rest_market_value = trim(rest_market_value)
'
'         If value_verif_code = "TX" Then value_verif_code = "TX - Property Tax Statement"
'         If value_verif_code = "PU" Then value_verif_code = "PU - Purchase Agreement"
'         If value_verif_code = "AP" Then value_verif_code = "AP - Appraisal"
'         If value_verif_code = "CO" Then value_verif_code = "CO - County Record"
'         If value_verif_code = "OT" Then value_verif_code = "OT - Other Document"
'         If value_verif_code = "NO" Then value_verif_code = "NO - No Ver Prvd"
'
'         rest_amt_owed = replace(rest_amt_owed, "_", "")
'         rest_amt_owed = trim(rest_amt_owed)
'
'         If amt_owed_verif_code = "MO" Then amt_owed_verif_code = "TI - Title/Deed"
'         If amt_owed_verif_code = "LN" Then amt_owed_verif_code = "CD - Contract for Deed"
'         If amt_owed_verif_code = "CD" Then amt_owed_verif_code = "CD - Contract for Deed"
'         If amt_owed_verif_code = "OT" Then amt_owed_verif_code = "OT - Other Document"
'         If amt_owed_verif_code = "NO" Then amt_owed_verif_code = "NO - No Ver Prvd"
'
'         rest_eff_date = replace(rest_eff_date, " ", "/")
'         If rest_eff_date = "__/__/__" Then rest_eff_date = ""
'
'         If rest_status = "1" Then rest_status = "1 - Home Residence"
'         If rest_status = "2" Then rest_status = "2 - For Sale, IV-E Rpymt Agmt"
'         If rest_status = "3" Then rest_status = "3 - Joint Owner, Unavailable"
'         If rest_status = "4" Then rest_status = "4 - Income Producing"
'         If rest_status = "5" Then rest_status = "5 - Future Residence"
'         If rest_status = "6" Then rest_status = "6 - Other"
'         If rest_status = "7" Then rest_status = "7 - For Sale, Unavailable"
'
'         rest_joint_yn = replace(rest_joint_yn, "_", "")
'         rest_ratio = replace(rest_ratio, "_", "")
'
'         repymt_agree_date = replace(repymt_agree_date, " ", "/")
'         If repymt_agree_date = "__/__/__" Then repymt_agree_date = ""
'
'         rest_update_date = replace(rest_update_date, " ", "/")
'
'         EMReadScreen panel_ref_numb, 1, 2, 73
'         panel_ref_numb = "0" & panel_ref_numb
'     End If
' end function
'
' function access_HEST_panel(access_type, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         Call navigate_to_MAXIS_screen("STAT", "HEST")
'
'         hest_col = 40
'         Do
'             EMReadScreen pers_paying, 2, 6, hest_col
'             If pers_paying <> "__" Then
'                 all_persons_paying = all_persons_paying & ", " & pers_paying
'             Else
'                 exit do
'             End If
'             hest_col = hest_col + 3
'         Loop until hest_col = 70
'         If left(all_persons_paying, 1) = "," Then all_persons_paying = right(all_persons_paying, len(all_persons_paying) - 2)
'
'         EMReadScreen choice_date, 8, 7, 40
'         EMReadScreen actual_initial_exp, 8, 8, 61
'
'         EMReadScreen retro_heat_ac_yn, 1, 13, 34
'         EMReadScreen retro_heat_ac_units, 2, 13, 42
'         EMReadScreen retro_heat_ac_amt, 6, 13, 49
'         EMReadScreen retro_electric_yn, 1, 14, 34
'         EMReadScreen retro_electric_units, 2, 14, 42
'         EMReadScreen retro_electric_amt, 6, 14, 49
'         EMReadScreen retro_phone_yn, 1, 15, 34
'         EMReadScreen retro_phone_units, 2, 15, 42
'         EMReadScreen retro_phone_amt, 6, 15, 49
'
'         EMReadScreen prosp_heat_ac_yn, 1, 13, 60
'         EMReadScreen prosp_heat_ac_units, 2, 13, 68
'         EMReadScreen prosp_heat_ac_amt, 6, 13, 75
'         EMReadScreen prosp_electric_yn, 1, 14, 60
'         EMReadScreen prosp_electric_units, 2, 14, 68
'         EMReadScreen prosp_electric_amt, 6, 14, 75
'         EMReadScreen prosp_phone_yn, 1, 15, 60
'         EMReadScreen prosp_phone_units, 2, 15, 68
'         EMReadScreen prosp_phone_amt, 6, 15, 75
'
'         choice_date = replace(choice_date, " ", "/")
'         If choice_date = "__/__/__" Then choice_date = ""
'         actual_initial_exp = trim(actual_initial_exp)
'         actual_initial_exp = replace(actual_initial_exp, "_", "")
'
'         retro_heat_ac_yn = replace(retro_heat_ac_yn, "_", "")
'         retro_heat_ac_units = replace(retro_heat_ac_units, "_", "")
'         retro_heat_ac_amt = trim(retro_heat_ac_amt)
'         retro_electric_yn = replace(retro_electric_yn, "_", "")
'         retro_electric_units = replace(retro_electric_units, "_", "")
'         retro_electric_amt = trim(retro_electric_amt)
'         retro_phone_yn = replace(retro_phone_yn, "_", "")
'         retro_phone_units = replace(retro_phone_units, "_", "")
'         retro_phone_amt = trim(retro_phone_amt)
'
'         prosp_heat_ac_yn = replace(prosp_heat_ac_yn, "_", "")
'         prosp_heat_ac_units = replace(prosp_heat_ac_units, "_", "")
'         prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
'         If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
'         prosp_electric_yn = replace(prosp_electric_yn, "_", "")
'         prosp_electric_units = replace(prosp_electric_units, "_", "")
'         prosp_electric_amt = trim(prosp_electric_amt)
'         If prosp_electric_amt = "" Then prosp_electric_amt = 0
'         prosp_phone_yn = replace(prosp_phone_yn, "_", "")
'         prosp_phone_units = replace(prosp_phone_units, "_", "")
'         prosp_phone_amt = trim(prosp_phone_amt)
'         If prosp_phone_amt = "" Then prosp_phone_amt = 0
'
'         total_utility_expense = 0
'         If prosp_heat_ac_yn = "Y" Then
'             total_utility_expense =  prosp_heat_ac_amt
'         ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
'             total_utility_expense =  prosp_electric_amt + prosp_phone_amt
'         ElseIf prosp_electric_yn = "Y" Then
'             total_utility_expense =  prosp_electric_amt
'         Elseif prosp_phone_yn = "Y" Then
'             total_utility_expense =  prosp_phone_amt
'         End If
'
'     End If
' end function
'
' function access_WREG_panel(access_type, notes_on_wreg, clt_fs_pwe, clt_wreg_status, clt_defer_fset, clt_orient_date, clt_sanc_begin_date, clt_numb_of_sanc, clt_sanc_reasons, clt_abawd_status, clt_banked_months, clt_GA_elig_basis, clt_GA_coop, abawd_counted_months, abawd_info_list, second_abawd_period, second_set_info_list)
'     access_type = UCase(access_type)
'     If access_type = "READ" Then
'         EMReadScreen clt_fs_pwe, 1, 6, 68
'         EMReadScreen clt_wreg_status, 2, 8, 50
'         EMReadScreen clt_defer_fset, 1, 8, 80
'         EMReadScreen clt_orient_date, 8, 9, 50
'         EMReadScreen clt_sanc_begin_date, 8, 10, 50
'         EMReadScreen clt_numb_of_sanc, 2, 11, 50
'         EMReadScreen clt_sanc_reasons, 2, 12, 50
'         EMReadScreen clt_abawd_status, 2, 13, 50
'         EMReadScreen clt_banked_months, 1, 14, 50
'         EMReadScreen clt_GA_elig_basis, 2, 15, 50
'         EMReadScreen clt_GA_coop, 2, 15, 78
'
'         EmWriteScreen "x", 13, 57
'         transmit
'         bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
'         bene_yr_row = 10
'         abawd_counted_months = 0
'         abawd_info_list = ""
'         second_abawd_period = 0
'         second_set_info_list = ""
'         month_count = 0
'         DO
'             'establishing variables for specific ABAWD counted month dates
'             If bene_mo_col = "19" then counted_date_month = "01"
'             If bene_mo_col = "23" then counted_date_month = "02"
'             If bene_mo_col = "27" then counted_date_month = "03"
'             If bene_mo_col = "31" then counted_date_month = "04"
'             If bene_mo_col = "35" then counted_date_month = "05"
'             If bene_mo_col = "39" then counted_date_month = "06"
'             If bene_mo_col = "43" then counted_date_month = "07"
'             If bene_mo_col = "47" then counted_date_month = "08"
'             If bene_mo_col = "51" then counted_date_month = "09"
'             If bene_mo_col = "55" then counted_date_month = "10"
'             If bene_mo_col = "59" then counted_date_month = "11"
'             If bene_mo_col = "63" then counted_date_month = "12"
'             'reading to see if a month is counted month or not
'             EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
'             'counting and checking for counted ABAWD months
'             IF is_counted_month = "X" or is_counted_month = "M" THEN
'                 EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
'                 abawd_counted_months_string = counted_date_month & "/" & counted_date_year
'                 abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
'                 abawd_counted_months = abawd_counted_months + 1				'adding counted months
'             END IF
'
'             'declaring & splitting the abawd months array
'             If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
'
'             'counting and checking for second set of ABAWD months
'             IF is_counted_month = "Y" or is_counted_month = "N" THEN
'                 EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
'                 second_abawd_period = second_abawd_period + 1				'adding counted months
'                 second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
'                 second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
'             END IF
'
'             'declaring & splitting the second set of abawd months array
'             If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
'
'             bene_mo_col = bene_mo_col - 4
'             IF bene_mo_col = 15 THEN
'                 bene_yr_row = bene_yr_row - 1
'                 bene_mo_col = 63
'             END IF
'             month_count = month_count + 1
'         LOOP until month_count = 36
'         PF3
'
'         clt_fs_pwe = replace(clt_fs_pwe, "_", "")
'         If clt_wreg_status = "03" Then clt_wreg_status = "03 - Unfit for Employment"
'         If clt_wreg_status = "04" Then clt_wreg_status = "04 - Resp for Care of Incapacitated Person"
'         If clt_wreg_status = "05" Then clt_wreg_status = "05 - Age 60 or Older"
'         If clt_wreg_status = "06" Then clt_wreg_status = "06 - Under Age 16"
'         If clt_wreg_status = "07" Then clt_wreg_status = "07 - Age 16-17, Living w/ Caregiver"
'         If clt_wreg_status = "08" Then clt_wreg_status = "08 - Resp for Care of Child under 6"
'         If clt_wreg_status = "09" Then clt_wreg_status = "09 - Empl 30 hrs/wk or Earnings of 30 hrs/wk"
'         If clt_wreg_status = "10" Then clt_wreg_status = "10 - Matching Grant Participant"
'         If clt_wreg_status = "11" Then clt_wreg_status = "11 - Receiving or Applied for UI"
'         If clt_wreg_status = "12" Then clt_wreg_status = "12 - Enrolled in School, Training, or Higher Ed"
'         If clt_wreg_status = "13" Then clt_wreg_status = "13 - Participating in CD Program"
'         If clt_wreg_status = "14" Then clt_wreg_status = "14 - Receiving MFIP"
'         If clt_wreg_status = "20" Then clt_wreg_status = "20 - Pending/Receiving DWP"
'         If clt_wreg_status = "15" Then clt_wreg_status = "15 - Age 16-17, NOT Living w/ Caregiver"
'         If clt_wreg_status = "16" Then clt_wreg_status = "16 - 50-59 Years Old"
'         If clt_wreg_status = "17" Then clt_wreg_status = "17 - Receiving RCA or GA"
'         If clt_wreg_status = "21" Then clt_wreg_status = "21 - Resp for Care of Child under 18"
'         If clt_wreg_status = "30" Then clt_wreg_status = "30 - Mandatory FSET Participant"
'         If clt_wreg_status = "02" Then clt_wreg_status = "02 - Fail to Cooperate with FSET"
'         If clt_wreg_status = "33" Then clt_wreg_status = "33 - Non-Coop being Referred"
'
'         clt_defer_fset = replace(clt_defer_fset, "_", "")
'         clt_orient_date = replace(clt_orient_date, " ", "/")
'         IF clt_orient_date = "__/__/__" Then clt_orient_date = ""
'
'         clt_sanc_begin_date = replace(clt_sanc_begin_date, " ", "/")
'         IF clt_sanc_begin_date = "__/01/__" Then clt_sanc_begin_date = ""
'         IF clt_numb_of_sanc = "01" Then clt_numb_of_sanc = "1st Sanction"
'         IF clt_numb_of_sanc = "02" Then clt_numb_of_sanc = "2nd Sanction"
'         IF clt_numb_of_sanc = "03" Then clt_numb_of_sanc = "3rd Sanction"
'         If clt_numb_of_sanc = "__" Then clt_numb_of_sanc = ""
'         If clt_sanc_reasons = "01" Then clt_sanc_reasons = "01 - Attend Orientation"
'         If clt_sanc_reasons = "02" Then clt_sanc_reasons = "02 - Develop Work Plan"
'         If clt_sanc_reasons = "03" Then clt_sanc_reasons = "03 - Follow Work Plan"
'         If clt_sanc_reasons = "__" Then clt_sanc_reasons = ""
'
'         If clt_abawd_status = "01" Then clt_abawd_status = "01 - Work Reg Exempt"
'         If clt_abawd_status = "02" Then clt_abawd_status = "02 - Under Age 18"
'         If clt_abawd_status = "03" Then clt_abawd_status = "03 - Age 50 or Over"
'         If clt_abawd_status = "04" Then clt_abawd_status = "04 - Caregiver of Minor Child"
'         If clt_abawd_status = "05" Then clt_abawd_status = "05 - Pregnant"
'         If clt_abawd_status = "06" Then clt_abawd_status = "06 - Employed Avg of 20 hrs/wk"
'         If clt_abawd_status = "07" Then clt_abawd_status = "07 - Work Experience Participant"
'         If clt_abawd_status = "08" Then clt_abawd_status = "08 - Other E&T Services"
'         If clt_abawd_status = "09" Then clt_abawd_status = "09 - Resides in a Waivered Area"
'         If clt_abawd_status = "10" Then clt_abawd_status = "10 - ABAWD Counted Month"
'         If clt_abawd_status = "11" Then clt_abawd_status = "11 - 2nd-3rd Month Period of Elig"
'         If clt_abawd_status = "12" Then clt_abawd_status = "12 - RCA or GA Recipient"
'         If clt_abawd_status = "13" Then clt_abawd_status = "13 - ABAWD Banked Months"
'         clt_banked_months = replace(clt_banked_months, "_", "")
'
'         If clt_GA_elig_basis = "04" Then clt_GA_elig_basis = "04 - Permanent Ill or Incap"
'         If clt_GA_elig_basis = "05" Then clt_GA_elig_basis = "05 - Temporary Ill or Incap"
'         If clt_GA_elig_basis = "06" Then clt_GA_elig_basis = "06 - Care of Ill or Incap Memb"
'         If clt_GA_elig_basis = "07" Then clt_GA_elig_basis = "07 - Requires Services in Residence"
'         If clt_GA_elig_basis = "09" Then clt_GA_elig_basis = "09 - Mentally Ill or Dev Disa"
'         If clt_GA_elig_basis = "10" Then clt_GA_elig_basis = "10 - SSI/RSDI Pending"
'         If clt_GA_elig_basis = "11" Then clt_GA_elig_basis = "11 - Appealing SSI/RSDI Denial"
'         If clt_GA_elig_basis = "12" Then clt_GA_elig_basis = "12 - Advanced Age"
'         If clt_GA_elig_basis = "13" Then clt_GA_elig_basis = "13 - Learning Disability"
'         If clt_GA_elig_basis = "17" Then clt_GA_elig_basis = "17 - Protect/Court Ordered"
'         If clt_GA_elig_basis = "20" Then clt_GA_elig_basis = "20 - Age 16 or 17 SS Approval"
'         If clt_GA_elig_basis = "25" Then clt_GA_elig_basis = "25 - Emancipated Minor"
'         If clt_GA_elig_basis = "28" Then clt_GA_elig_basis = "28 - Unemployable"
'         If clt_GA_elig_basis = "29" Then clt_GA_elig_basis = "29 - Displaced Hmkr (FT Student)"
'         If clt_GA_elig_basis = "30" Then clt_GA_elig_basis = "30 - Minor w/ Adult Unrelated"
'         If clt_GA_elig_basis = "32" Then clt_GA_elig_basis = "32 - Adult ESL/Adult HS"
'         If clt_GA_elig_basis = "99" Then clt_GA_elig_basis = "99 - No Elig Basis"
'         If clt_GA_elig_basis = "__" Then clt_GA_elig_basis = ""
'
'         If clt_GA_coop = "01" Then clt_GA_coop = "01 - Cooperating"
'         If clt_GA_coop = "03" Then clt_GA_coop = "03 - Failed to Coop"
'         If clt_GA_coop = "__" Then clt_GA_coop = ""
'     end If
' end function

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
			  ' Text 10, 135, 75, 10, "Residence Address:"
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
			  ' Text 20, 145, 110, 10, "resi_line_one"
			  ' Text 20, 155, 115, 10, "resi_city & ,  & resi_state &   & resi_zip"
			  ' Text 205, 135, 75, 10, "Mailing Address:"
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
			  ' Text 210, 145, 110, 10, "mail_line_one"
			  ' Text 210, 155, 120, 10, "mail_city & ,  & mail_state &   & mail_zip"
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
		        ' If snap_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a SNAP Six-Month report indicate the status of the SR process (eg. N, I, or U)."
		        Call validate_footer_month_entry(snap_sr_mo, snap_sr_yr, err_msg, "* SNAP SR MONTH")
		        program_indicated = TRUE
		    End If
		    If hc_sr_yn = "Yes" Then
		        ' If hc_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a HC Six-Month report indicate the status of the SR process (eg. N, I, or U)."
		        Call validate_footer_month_entry(hc_sr_mo, hc_sr_yr, err_msg, "* HC SR MONTH")
		        program_indicated = TRUE
		    End If
		    If grh_sr_yn = "Yes" Then
		        ' If grh_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a GRH Six-Month report indicate the status of the SR process (eg. N, I, or U)."
		        Call validate_footer_month_entry(grh_sr_mo, grh_sr_yr, err_msg, "* GRH SR MONTH")
		        program_indicated = TRUE
		    End If

			If ButtonPressed = -1 Then
				err_msg = ""
			    If client_on_csr_form = "Select or Type" OR trim(client_on_csr_form) = "" Then err_msg = err_msg & vbNewLine & "* Indicate who is listed on the CSR form in the person infromation, or if this is blank, select that the person information is missing."

			    If program_indicated = FALSE Then err_msg = err_msg & vbNewLine & "* Select the program(s) that the CSR form is processing. (None of the programs are indicated to have an SR due.)"

			    ' If residence_address_match_yn = "Does the residence address match?" Then err_msg = err_msg & vbNewLine & "* Indicate information about the residence address provided on the CSR form."
			    ' If mailing_address_match_yn = "Does the mailing address match?" Then err_msg = err_msg & vbNewLine & "* Indicate information abobut the mailing address provided on the CSR form."
				' If residence_address_match_yn = "No - New Address Entered" AND new_resi_addr_entered = FALSE Then err_msg = err_msg & vbNewLine & "* The option 'No - New Address Endered' for the residence address can only be updated by the script."
				' If mailing_address_match_yn = "No - New Address Entered" AND new_mail_addr_entered = FALSE Then err_msg = err_msg & vbNewLine & "* The option 'No - New Address Endered' for the Mailing address can only be updated by the script."
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
		  ' GroupBox 5, 180, 415, grp_len, "Household Comp"
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
		  ' y_pos = y_pos + 25
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
		' If apply_for_ma = "Yes" Then
		'     dlg_len = dlg_len + (UBound(NEW_MA_REQUEST_ARRAY, 2) + 1) * 20
		'     q_4_grp_len = 35 + UBound(NEW_MA_REQUEST_ARRAY, 2) * 20
		' End If
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
		  ' DropListBox 285, 25, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", apply_for_ma
		  ' CheckBox 430, 30, 75, 10, "Q4 Deailts left Blank", q_4_details_blank_checkbox
		  DropListBox 285, 25, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", apply_for_ma
		  ButtonGroup ButtonPressed
			PushButton 540, 30, 50, 10, "Add Another", add_memb_btn
		  ' If apply_for_ma = "Yes" Then
		  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
			  If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
				  Text 35, y_pos + 5, 105, 10, "Select the Member requesting:"
				  ComboBox 145, y_pos, 195, 45, all_the_clients, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If y_pos = 45 Then y_pos = y_pos + 5
		  ' End If

		  GroupBox 15, y_pos + 5, 585, q_5_grp_len, "Q5. Is anyone self-employed or does anyone expect to be self-employed?"
		  ' DropListBox 265, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_self_employed
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q5 Deailts left Blank", q_5_details_blank_checkbox
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
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_busi)
				  ' CheckBox 530, y_pos, 40, 10, "Detail", ALL_INCOME_ARRAY(update_checkbox, each_busi)
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
		  ' DropListBox 230, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_start_working
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q6 Deailts left Blank", q_6_details_blank_checkbox
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
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_job)
				  ' CheckBox 530, y_pos, 40, 10, "Detail", ALL_INCOME_ARRAY(update_checkbox, each_job)
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
		  ' DropListBox 335, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_other_income
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q7 Deailts left Blank", q_7_details_blank_checkbox
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
				  ComboBox 165, y_pos, 110, 45, unea_type_list, NEW_UNEARNED_ARRAY(unearned_source, each_unea)   'unea_type
				  EditBox 280, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)    'unea_start_date
				  EditBox 335, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_amount, each_unea)    'unea_amount
				  DropListBox 390, y_pos, 90, 45, "Select One..."+chr(9)+"4 - Weekly"+chr(9)+"3 - Biweekly"+chr(9)+"2 - Semi Monthly"+chr(9)+"1 - Monthly", NEW_UNEARNED_ARRAY(unearned_freq, each_unea) 'unea_frequency
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_unea)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_INCOME_ARRAY(update_checkbox, each_unea)
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
			' PushButton 20, y_pos + 2, 200, 13, "Why do I have to answer these if in is not HC?", why_answer_btn
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

			' q_4_details_blank_checkbox = checked
			' q_5_details_blank_checkbox = checked
			' q_6_details_blank_checkbox = checked
			' q_7_details_blank_checkbox = checked
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
			new_item = UBound(ALL_INCOME_ARRAY, 2) + 1
			ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, new_unearned_counter)
			NEW_UNEARNED_ARRAY(unearned_type, new_unearned_counter) = "UNEA"
			NEW_UNEARNED_ARRAY(unearned_prog_list, new_unearned_counter) = "MA"
			new_unearned_counter = new_unearned_counter + 1
		End If
		' If ButtonPressed = why_answer_btn Then
		' 	explain_text = "This case may not have MA, MSP, or any HC active and you may have indicated that it is only for a SNAP Review, HOWEVER" & vbCr & vbCr
		' 	explain_text = explain_text & "The form that was sent to the client STILL has these questions listed on it." & vbCr
		' 	explain_text = explain_text & "We need to be looking at all information that the client reported, anything entered here may impact the benefits because it is now 'known to the agency'." & vbCr & vbCr
		' 	explain_text = explain_text & "Though the client is not required to answer these questions, we are still required to review the entire form."
		' 	' explain_text = explain_text & ""
		' 	why_answer_when_not_HC_msg = MsgBOx(explain_text, vbInformation + vbOKonly, "No HC on the case")
		' End If

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
			' If q_4_details_blank_checkbox = unchecked Then
			If InStr(apply_for_ma, "details listed below") <> 0 Then
				For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
					If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Enter or Select Member" Then
						q_4_details_entered = TRUE
					End If
				Next
				' If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No details of a person requesting MA for someone not getting coverage now (Question 4). Either enter information about which members are requesting MA coverage or check the box to indicate this portion of the form was left blank."
				' If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 4 - Someone getting MA coverage. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_4_details_entered = TRUE
			End If
			' If q_5_details_blank_checkbox = unchecked Then
			If InStr(ma_self_employed, "details listed below") <> 0 Then
				For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
						q_5_details_entered = TRUE
					End If
				Next
				' If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No Self Employment information has been entered (Question 5). Either enter BUSI details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 5 - Self-Employment Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_5_details_entered = TRUE
			End If
			' If q_6_details_blank_checkbox = unchecked Then
			If InStr(ma_start_working, "details listed below") <> 0 Then
				For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
						q_6_details_entered = TRUE
					End If
				Next
				' If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Job information has been entered (Question 6). Either enter JOBS details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 6 - Job Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_6_details_entered = TRUE
			End If
			' If q_7_details_blank_checkbox = unchecked Then
			If InStr(ma_other_income, "details listed below") <> 0 Then
				For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
					If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
						q_7_details_entered = TRUE
					End If
				Next
				' If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 7). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 7 - Unearned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
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
				' MsgBox ALL_ASSETS_ARRAY(category_const, assets_on_case)
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
		  ' DropListBox 330, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_liquid_assets
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q9 Deailts left Blank", q_9_details_blank_checkbox
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
			  ComboBox 165, y_pos, 115, 40, account_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)'liquid_asst_type
			  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)'liquid_asset_name
			  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)    'new_checkbox
			  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)'update_checkbox
			  y_pos = y_pos + 20
			End If
		  Next
		  If first_account = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no ACCT information has been added."
			y_pos = y_pos + 10
		  End If

		  y_pos = y_pos +10
		  GroupBox 15, y_pos + 5, 585, q_10_grp_len, "Q10. Does anyone own or co-own securities or other assets?"
		  ' DropListBox 295, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_security_assets
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q10 Deailts left Blank", q_10_details_blank_checkbox
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
				ComboBox 165, y_pos, 115, 40, security_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'security_asset_type
				EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)   'security_asset_name
				' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
				y_pos = y_pos + 20
			End If
		  Next
		  If first_secu = TRUE Then
			  Text 30, y_pos, 250, 10, "CSR form - no SECU information has been added."
			  y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_11_grp_len, "Q11. Does anyone own a vehicle?"
		  ' DropListBox 250, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_vehicle
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q11 Deailts left Blank", q_11_details_blank_checkbox
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
				  ComboBox 165, y_pos, 115, 40, cars_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'vehicle_asset_type
				  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_year_make_model, each_asset)  'vehicle_asset_name
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_car = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no CARS information has been added."
			y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_12_grp_len, "Q12. Does anyone own or co-own any real estate?"
		  ' DropListBox 280, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_real_assets
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q12 Deailts left Blank", q_12_details_blank_checkbox
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
				  ComboBox 320, y_pos, 150, 40, rest_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)     'property_asset_type
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
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

		' MsgBox ButtonPressed & " - 1 - "
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

			' q_9_details_blank_checkbox = checked
			' q_10_details_blank_checkbox = checked
			' q_11_details_blank_checkbox = checked
			' q_12_details_blank_checkbox = checked
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
			' If q_9_details_blank_checkbox = unchecked Then
			If InStr(ma_liquid_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" AND NEW_ASSET_ARRAY(ma_request_client, each_asset) <> "Enter or Select Member" Then
						q_9_details_entered = TRUE
					End If
				Next
				' If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No details of a person requesting MA for someone not getting coverage now (Question 9). Either enter information about which members are requesting MA coverage or check the box to indicate this portion of the form was left blank."
				' If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 9 - Liquid Assets. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_9_details_entered = TRUE
			End If
			' If q_10_details_blank_checkbox = unchecked Then
			If InStr(ma_security_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_10_details_entered = TRUE
					End If
				Next
				' If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No Self Employment information has been entered (Question 10). Either enter BUSI details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 10 - Securities. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_10_details_entered = TRUE
			End If
			' If q_11_details_blank_checkbox = unchecked Then
			If InStr(ma_vehicle, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_11_details_entered = TRUE
					End If
				Next
				' If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Job information has been entered (Question 11). Either enter JOBS details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 11 - Vehicles. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_11_details_entered = TRUE
			End If
			' If q_12_details_blank_checkbox = unchecked Then
			If InStr(ma_real_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_12_details_entered = TRUE
					End If
				Next
				' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 12 - Real Estate. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
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
			' MsgBox ButtonPressed & " - 2 - "
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
		  ' DropListBox 160, 20, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", ma_other_changes
		  ' CheckBox 30, 60, 300, 10, "Check here if client left the changes to report field on the form blank.", changes_reported_blank_checkbox
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
	' MsgBox "Q 13 Cleared - " & csr_dlg_q_13_cleared
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
		' dlg_len = dlg_len + UBound(NEW_EARNED_ARRAY, 2) * 20
		' dlg_len = dlg_len + UBound(NEW_UNEARNED_ARRAY, 2) * 20
		' dlg_len = dlg_len + UBound(NEW_CHILD_SUPPORT_ARRAY, 2) * 20
		' q_15_grp_len = 50
		' q_16_grp_len = 45 + UBound(NEW_EARNED_ARRAY, 2) * 20
		' q_17_grp_len = 45 + UBound(NEW_UNEARNED_ARRAY, 2) * 20
		' q_18_grp_len = 45 + UBound(NEW_CHILD_SUPPORT_ARRAY, 2) * 20

		y_pos = 95

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 615, dlg_len, "SNAP CSR Question Details"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 15 - 19:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 15 - 19 are not required.", all_questions_15_19_blank
		  GroupBox 10, 30, 600, q_15_grp_len, "Q15. Has your household moved since your last application or in the past six months?"
		  ' DropListBox 305, 25, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_fifteen_form_answer
		  DropListBox 305, 25, 150, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Not Required", quest_fifteen_form_answer
		  Text 25, 45, 105, 10, "New Rent or Mortgage Amount:"
		  EditBox 130, 40, 65, 15, new_rent_or_mortgage_amount
		  CheckBox 220, 45, 50, 10, "Heat/AC", heat_ac_checkbox
		  CheckBox 275, 45, 50, 10, "Electricity", electricity_checkbox
		  CheckBox 345, 45, 50, 10, "Telephone", telephone_checkbox
		  ' Text 400, 45, 80, 10, "Did client attach proof?"
		  ' DropListBox 480, 40, 125, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", shel_proof_provided
		  GroupBox 10, 70, 490, q_16_grp_len, "Q16 Has there been a change in EARNED INCOME?"
		  ' DropListBox 190, 65, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_sixteen_form_answer
		  ' CheckBox 310, 70, 85, 10, "Q16 Deailts left Blank", q_16_details_blank_checkbox
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
		  ' DropListBox 205, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_seventeen_form_answer
		  ' CheckBox 310, y_pos, 85, 10, "Q17 Deailts left Blank", q_17_details_blank_checkbox
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
		  ' DropListBox 190, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_eighteen_form_answer
		  ' CheckBox 310, y_pos, 85, 10, "Q18 Deailts left Blank", q_18_details_blank_checkbox
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
		  ' y_pos = y_pos + 15
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

			' q_16_details_blank_checkbox = checked
			' q_17_details_blank_checkbox = checked
			' q_18_details_blank_checkbox = checked
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

			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_15_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 15 - Shelter and Utilities Expenses. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_15_details_entered = TRUE
		End If

		If InStr(quest_sixteen_form_answer, "details listed below") <> 0 Then
			For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
				If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
					q_16_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_16_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 16 - Earned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_16_details_entered = TRUE
		End If

		If InStr(quest_seventeen_form_answer, "details listed below") <> 0 Then
			For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
				If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
					q_17_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_17_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 17 - Unearned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_17_details_entered = TRUE
		End If

		If InStr(quest_eighteen_form_answer, "details listed below") <> 0 Then
			For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
				If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
					q_18_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 18 - Child Support. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
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
		' MsgBox show_two & vbNewLine & "line 1480"
		' Loop until leave_ma_questions = TRUE
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
		' If client_dated_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the form has been dated correctly by selecting 'yes' or 'no'."

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
	  ' Text 20, 160, 115, 10, "your home in the past six months?"
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
			  ' Text 25, 205, 155, 10, "Date of Change:"
			  ' Text 25, 215, 155, 10, "other"
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
			  ' Text 25, 205, 155, 10, "Date of Change:"
			  ' Text 25, 215, 155, 10, "other"
			  pers_list_count = pers_list_count + 1
			  y_pos_1 = y_pos_1 + 40
		  End If
	  Next


	  GroupBox 190, 5, 185, 340, "Page 2"
	  Text 195, 20, 135, 20, "4. Do you want to apply for someone who is not getting coverage now?"
	  ' Text 205, 30, 125, 10, "is not getting coverage now?"
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
	  ' Text 205, 60, 125, 10, "anyone expect to be self-employed?"
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
	  ' Text 205, 90, 125, 10, "expect to start working?"
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
	  ' Text 205, 120, 130, 10, "anyone expect to get money from "
	  ' Text 205, 130, 115, 10, "sources other than work?"
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
	  ' Text 390, 50, 125, 10, "other assets?"
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
	  ' Text 390, 100, 130, 10, "or any real estate?"
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

		  ' Text 570, y_pos_4, 150, 10, dlg_text

	  End If

	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 20, "16. Has anyone had a change in their income from work?"
	  ' Text 575, 75, 125, 10, ""
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
	  ' Text 575, 105, 140, 10, "than $50 per month from income sources"
	  ' Text 575, 115, 160, 10, "other than work or a change in unearned income?"
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
	  ' Text 575, 145, 140, 10, "court-ordered child or medical "
	  ' Text 575, 155, 160, 10, "support payments?"
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
	  ' Text 575, 185, 140, 10, "for an average of 80 hours each month"
	  ' Text 575, 195, 160, 10, "during the past six months?"
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
	' MsgBox show_csr_dlg_q_15_19
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
	local_work_save_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_work_save_path) = True then
			.DeleteFile(local_work_save_path)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

		If .FileExists(local_work_save_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(local_work_save_path, ForWriting, true)

			'Write the contents of the text file
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
			' objTextStream.WriteLine "15 - PROF - " & shel_proof_provided
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
			' objTextStream.WriteLine
			' objTextStream.WriteLine
			' objTextStream.WriteLine
			' objTextStream.WriteLine

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

				known_memb = 0
				new_hh_memb = 0
				each_new_memb = 0
				each_job = 0
				each_unea = 0
				each_asset = 0
				the_cs = 0
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
						' MsgBox Mid(text_line, 6, 13)
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

						' objTextStream.WriteLine "NEW_MEMBERS_ARRAY - " &
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

						' objTextStream.WriteLine "NEW_MA_REQUEST_ARRAY - " &
						ReDim Preserve NEW_MA_REQUEST_ARRAY(ma_request_notes, each_new_memb)
						NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) = array_info(0)
						each_new_memb = each_new_memb + 1
					End If
					If left(text_line, 16) = "NEW_EARNED_ARRAY" Then
						array_info = Mid(text_line, 20)
						array_info = split(array_info, "~")

						' objTextStream.WriteLine "NEW_EARNED_ARRAY - " &
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

						' objTextStream.WriteLine "NEW_UNEARNED_ARRAY - " &
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

						' objTextStream.WriteLine "NEW_ASSET_ARRAY - " &
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

						' objTextStream.WriteLine "NEW_CHILD_SUPPORT_ARRAY - " &
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

Const owner_name                = 00
Const category_const            = 01
Const type_const                = 02
Const name_const                = 03
Const amount_const              = 04
Const start_date_const          = 05
Const end_date_const            = 06
Const verif_const               = 07
Const pay_amt_const             = 08
Const hours_const               = 09
Const update_date_const         = 10
Const seasonal_yn               = 11
Const frequency_const           = 12
Const make_const                = 13
Const model_const               = 14
Const year_const                = 15
Const make_model_yr             = 16
Const address_const             = 17
Const cash_amt_const            = 18        'cash panel
Const cash_verif_const          = 19
Const snap_amt_const            = 20
Const snap_verif_const          = 21
Const hc_amt_const              = 22
Const hc_verif_const            = 23
Const busi_cash_net_prosp       = 24        'busi panel
Const busi_cash_net_retro       = 25
Const busi_cash_gross_retro     = 26
Const busi_cash_expense_retro   = 27
Const busi_cash_gross_prosp     = 28
Const busi_cash_expense_prosp   = 29
Const busi_cash_income_verif    = 30
Const busi_cash_expense_verif   = 31
Const busi_snap_net_prosp       = 32
Const busi_snap_net_retro       = 33
Const busi_snap_gross_retro     = 34
Const busi_snap_expense_retro   = 35
Const busi_snap_gross_prosp     = 36
Const busi_snap_expense_prosp   = 37
Const busi_snap_income_verif    = 38
Const busi_snap_expense_verif   = 39
Const busi_hc_a_net_prosp       = 40
Const busi_hc_a_gross_prosp     = 41
Const busi_hc_a_expense_prosp   = 42
Const busi_hc_a_income_verif    = 43
Const busi_hc_a_expense_verif   = 44
Const busi_hc_b_net_prosp       = 45
Const busi_hc_b_gross_prosp     = 46
Const busi_hc_b_expense_prosp   = 47
Const busi_hc_b_income_verif    = 48
Const busi_hc_b_expense_verif   = 49
Const busi_se_method            = 50
Const busi_se_method_date       = 51
Const rptd_hours_const          = 52
Const min_wg_hours_const        = 53
Const claim_nbr_const           = 54        'unea panel
Const cola_disregard_amt        = 55
Const id_number_const           = 56
Const panel_instance            = 57
Const owner_ref_const           = 58
Const verif_checkbox_const      = 59
Const verif_time_const          = 60
Const verif_added_const         = 61
Const item_notes_const          = 62
Const balance_date_const        = 63
Const withdraw_penalty_const    = 64
Const withdraw_yn_const         = 65
Const withdraw_verif_const      = 66
Const count_cash_const          = 67
Const count_snap_const          = 68
Const count_hc_const            = 69
Const count_grh_const           = 70
Const count_ive_const           = 71
Const joint_own_const           = 72
Const share_ratio_const         = 73
Const next_interst_const        = 74
Const face_value_const          = 75
Const trade_in_const            = 76
Const loan_const                = 77
Const source_const              = 78
Const owed_amt_const            = 79
Const owed_verif_const          = 80
Const owed_date_const           = 81
Const cars_use_const            = 82
Const hc_benefit_const          = 83
Const market_value_const        = 84
Const value_verif_const         = 85
Const rest_prop_status_const    = 86
Const rest_repymt_date_const    = 87

Const jobs_hrly_wage            = 88
Const retro_income_amount       = 89
Const retro_income_hours        = 90
Const snap_pic_frequency        = 91
Const snap_pic_hours_per_pay    = 92
Const snap_pic_income_per_pay   = 93
Const snap_pic_monthly_income   = 94
Const grh_pic_frequency         = 95
Const grh_pic_income_per_pay    = 96
Const grh_pic_monthly_income    = 97
Const jobs_subsidy              = 98

Const new_checkbox              = 99
Const update_checkbox           = 100

Const faci_ref_numb                 = 00
Const faci_instance                 = 01
Const faci_member                   = 02
Const faci_name                     = 03
Const faci_vendor_number            = 04
Const faci_type                     = 05
Const faci_FS_elig                  = 06
Const faci_FS_type                  = 07
Const faci_waiver_type              = 08
Const faci_ltc_inelig_reason        = 09
Const faci_inelig_begin_date        = 10
Const faci_inelig_end_date          = 11
Const faci_anticipated_out_date     = 12
Const faci_GRH_plan_required        = 13
Const faci_GRH_plan_verif           = 14
Const faci_cty_app_place            = 15
Const faci_approval_cty_name        = 16
Const faci_GRH_DOC_amount           = 17
Const faci_GRH_postpay              = 18
Const faci_stay_one_rate            = 19
Const faci_stay_one_date_in         = 20
Const faci_stay_one_date_out        = 21
Const faci_stay_two_rate            = 22
Const faci_stay_two_date_in         = 23
Const faci_stay_two_date_out        = 24
Const faci_stay_three_rate          = 25
Const faci_stay_three_date_in       = 26
Const faci_stay_three_date_out      = 27
Const faci_stay_four_rate           = 28
Const faci_stay_four_date_in        = 29
Const faci_stay_four_date_out       = 30
Const faci_stay_five_rate           = 31
Const faci_stay_five_date_in        = 32
Const faci_stay_five_date_out       = 33
Const faci_verif_checkbox           = 34
Const faci_verif_added              = 35
Const faci_notes                    = 36

Dim ALL_INCOME_ARRAY()
ReDim ALL_INCOME_ARRAY(update_checkbox, 0)

Dim ALL_ASSETS_ARRAY()
ReDim ALL_ASSETS_ARRAY(update_checkbox, 0)

Dim FACILITIES_ARRAY()
ReDim FACILITIES_ARRAY(faci_notes, 0)

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
' const asset_
' const asset_
' const asset_
' const asset_
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

unea_type_list = "Type or Select"
unea_type_list = unea_type_list+chr(9)+"01 - RSDI, Disa"
unea_type_list = unea_type_list+chr(9)+"02 - RSDI, No Disa"
unea_type_list = unea_type_list+chr(9)+"03 - SSI"
unea_type_list = unea_type_list+chr(9)+"06 - Non-MN PA"
unea_type_list = unea_type_list+chr(9)+"11 - VA Disability"
unea_type_list = unea_type_list+chr(9)+"12 - VA Pension"
unea_type_list = unea_type_list+chr(9)+"13 - VA Other"
unea_type_list = unea_type_list+chr(9)+"38 - VA Aid & Attendance"
unea_type_list = unea_type_list+chr(9)+"14 - Unemployment Insurance"
unea_type_list = unea_type_list+chr(9)+"15 - Worker's Comp"
unea_type_list = unea_type_list+chr(9)+"16 - Railroad Retirement"
unea_type_list = unea_type_list+chr(9)+"17 - Other Retirement"
unea_type_list = unea_type_list+chr(9)+"18 - Military Enrirlement"
unea_type_list = unea_type_list+chr(9)+"19 - FC Child req FS"
unea_type_list = unea_type_list+chr(9)+"20 - FC Child not req FS"
unea_type_list = unea_type_list+chr(9)+"21 - FC Adult req FS"
unea_type_list = unea_type_list+chr(9)+"22 - FC Adult not req FS"
unea_type_list = unea_type_list+chr(9)+"23 - Dividends"
unea_type_list = unea_type_list+chr(9)+"24 - Interest"
unea_type_list = unea_type_list+chr(9)+"25 - Cnt gifts/prizes"
unea_type_list = unea_type_list+chr(9)+"26 - Strike Benefits"
unea_type_list = unea_type_list+chr(9)+"27 - Contract for Deed"
unea_type_list = unea_type_list+chr(9)+"28 - Illegal Income"
unea_type_list = unea_type_list+chr(9)+"29 - Other Countable"
unea_type_list = unea_type_list+chr(9)+"30 - Infrequent"
unea_type_list = unea_type_list+chr(9)+"31 - Other - FS Only"
unea_type_list = unea_type_list+chr(9)+"08 - Direct Child Support"
unea_type_list = unea_type_list+chr(9)+"35 - Direct Spousal Support"
unea_type_list = unea_type_list+chr(9)+"36 - Disbursed Child Support"
unea_type_list = unea_type_list+chr(9)+"37 - Disbursed Spousal Support"
unea_type_list = unea_type_list+chr(9)+"39 - Disbursed CS Arrears"
unea_type_list = unea_type_list+chr(9)+"40 - Disbursed Spsl Sup Arrears"
unea_type_list = unea_type_list+chr(9)+"43 - Disbursed Excess CS"
unea_type_list = unea_type_list+chr(9)+"44 - MSA - Excess Income for SSI"
unea_type_list = unea_type_list+chr(9)+"47 - Tribal Income"
unea_type_list = unea_type_list+chr(9)+"48 - Trust Income"
unea_type_list = unea_type_list+chr(9)+"49 - Non-Recurring"

account_list = "Select or Type"
account_list = account_list+chr(9)+"Cash"
account_list = account_list+chr(9)+"SV - Savings"
account_list = account_list+chr(9)+"CK - Checking"
account_list = account_list+chr(9)+"CE - Certificate of Deposit"
account_list = account_list+chr(9)+"MM - Money Market"
account_list = account_list+chr(9)+"DC - Debit Card"
account_list = account_list+chr(9)+"KO - Keogh Account"
account_list = account_list+chr(9)+"FT - Fed Thrift Savings Plan"
account_list = account_list+chr(9)+"SL - State & Local Govt"
account_list = account_list+chr(9)+"RA - Employee Ret Annuities"
account_list = account_list+chr(9)+"NP - Non-Profit Emmployee Ret"
account_list = account_list+chr(9)+"IR - Indiv Ret Acct"
account_list = account_list+chr(9)+"RH - Roth IRA"
account_list = account_list+chr(9)+"FR - Ret Plan for Employers"
account_list = account_list+chr(9)+"CT - Corp Ret Trust"
account_list = account_list+chr(9)+"RT - Other Ret Fund"
account_list = account_list+chr(9)+"QT - Qualified Tuition (529)"
account_list = account_list+chr(9)+"CA - Coverdell SV (530)"
account_list = account_list+chr(9)+"OE - Other Educational"
account_list = account_list+chr(9)+"OT - Other"

security_list = "Select or Type"
security_list = security_list+chr(9)+"LI - Life Insurance"
security_list = security_list+chr(9)+"ST - Stocks"
security_list = security_list+chr(9)+"BO - Bonds"
security_list = security_list+chr(9)+"CD - Ctrct for Deed"
security_list = security_list+chr(9)+"MO - Mortgage Note"
security_list = security_list+chr(9)+"AN - Annuity"
security_list = security_list+chr(9)+"OT - Other"

cars_list = "Select or Type"
cars_list = cars_list+chr(9)+"1 - Car"
cars_list = cars_list+chr(9)+"2 - Truck"
cars_list = cars_list+chr(9)+"3 - Van"
cars_list = cars_list+chr(9)+"4 - Camper"
cars_list = cars_list+chr(9)+"5 - Motorcycle"
cars_list = cars_list+chr(9)+"6 - Trailer"
cars_list = cars_list+chr(9)+"7 - Other"

rest_list = "Select or Type"
rest_list = rest_list+chr(9)+"1 - House"
rest_list = rest_list+chr(9)+"2 - Land"
rest_list = rest_list+chr(9)+"3 - Buildings"
rest_list = rest_list+chr(9)+"4 - Mobile Home"
rest_list = rest_list+chr(9)+"5 - Life Estate"
rest_list = rest_list+chr(9)+"6 - Other"

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
	' msgbox "reading things"
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	Dim HH_member_array()
	ReDim HH_member_array(0)

	hh_count = 0
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadScreen access_denied_check, 13, 24, 2
		'MsgBox access_denied_check
		If access_denied_check <> "ACCESS DENIED" Then
			ReDim Preserve HH_member_array(hh_count)
			HH_member_array(hh_count) = ref_nbr
			hh_count = hh_count + 1
		End If
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

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

	'check to see if there is an adult on MA'
	Call navigate_to_MAXIS_screen("STAT", "REVW")

	grh_sr = FALSE
	snap_sr = FALSE
	hc_sr = FALSE
	'Read for GRH
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

	'Read for SNAP
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

	'Read for MA
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

	Call generate_client_list(all_the_clients, "Select or Type")
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)
	full_hh_list = Split(list_for_array, chr(9))

	Call back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "MEMB")


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

	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, living_situation_status, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, curr_phone_one, curr_phone_two, curr_phone_three, curr_phone_type_one, curr_phone_type_two, curr_phone_type_three)
End If

new_memb_counter = 0
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

next_page_ma_btn = 1100
previous_page_btn = 1200
continue_btn = 1300

new_earned_counter = 0
new_unearned_counter = 0
new_asset_counter = 0



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
' Const wdDialogFilePrint = 88
' Const end_of_doc = 6
objWord.Caption = "CSR Form Details - CASE #" & MAXIS_case_number
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "CSR Information"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR
objSelection.TypeText "Date Completed: " & date & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR

If snap_sr_yn = "Yes" Then objSelection.TypeText "SNAP SR for " & snap_sr_mo & "/" & snap_sr_yr & vbCr
If hc_sr_yn = "Yes" Then objSelection.TypeText "HC SR for " & hc_sr_mo & "/" & hc_sr_yr & vbCr
If grh_sr_yn = "Yes" Then objSelection.TypeText "GRH SR for " & grh_sr_mo & "/" & grh_sr_yr & vbCr
objSelection.Font.Size = "11"

objSelection.TypeText "CSR Questions:" & vbCr
objSelection.TypeText "Q 1. Name and Address" & vbCr
objSelection.TypeText chr(9) & "Name: " & client_on_csr_form & vbCr
If new_resi_addr_entered = TRUE Then objSelection.TypeText chr(9) & "Address: " & new_resi_one & " " & new_resi_city & ", " & new_resi_state & " " & new_resi_zip & vbCr
If new_resi_addr_entered = FALSE Then objSelection.TypeText chr(9) & "Address: " & resi_line_one & " " & resi_line_two & " " & resi_city & ", " & resi_state & " " & resi_zip & vbCr

If new_mail_addr_entered = TRUE Then objSelection.TypeText chr(9) & "Mailing Address: " & new_mail_one & " " & new_mail_city & ", " & new_mail_state & " " & new_mail_zip & vbCr
If new_mail_addr_entered = FALSE Then objSelection.TypeText chr(9) & "Mailing Address: " & mail_line_one & " " & mail_line_two & " " & mail_city & ", " & mail_state & " " & mail_zip & vbCr
If homeless_status = "Yes" Then  objSelection.TypeText "Reports Homeless" & vbCr

objSelection.TypeText "Q 2. Has anyone moved in or out of your home in the past six months?" & vbCr
objSelection.TypeText chr(9) & quest_two_move_in_out & vbCr
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
table_counter = 1
objSelection.TypeText "Q 4. Do you want to apply for someone who is not getting coverage now?" & vbCr
objSelection.TypeText chr(9) & apply_for_ma & vbCr
q_4_count = 0
For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
	If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
		found_q_4_det = TRUE
		q_4_count = q_4_count + 1
	End If
Next
If found_q_4_det = TRUE Then
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_4_count + 1, 1
	set objQ4Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
	objQ4Table.Cell(1, 1).Range.Text = "Members "
	For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
		If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
			objQ4Table.Cell(tbl_row, 1).Range.Text = NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
			tbl_row = tbl_row + 1
		End If
	Next
	objQ4Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_5_count, 4
	set objQ5Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ5Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_6_count + 1, 6
	set objQ6Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ6Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_7_count + 1, 5
	set objQ7Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ7Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_9_count + 1, 3
	set objQ9Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ9Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_10_count + 1, 3
	set objQ10Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ10Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_11_count + 1, 3
	set objQ11Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ11Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_12_count + 1, 3
	set objQ12Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ12Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_16_count + 1, 6
	set objQ16Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ16Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_17_count + 1, 5
	set objQ17Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ17Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
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
	Set objRange = objSelection.Range
	objDoc.Tables.Add objRange, q_18_count + 1, 3
	set objQ18Table = objDoc.Tables(table_counter)
	table_counter = table_counter + 1
	tbl_row = 2
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
	objQ18Table.AutoFormat(16)
	objSelection.EndKey end_of_doc
	objSelection.TypeParagraph()
End If
objSelection.TypeText "Q 19. Did you work 20 hours each week, for an average of 80 hours each month during the past six months?" & vbCr
objSelection.TypeText chr(9) & quest_nineteen_form_answer & vbCr

objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE
objSelection.TypeText "Verbal Signature accepted on " & csr_form_date

t_drive = "\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team"
file_safe_date = replace(date, "/", "-")
word_doc_path = t_drive & "\Eligibility Support\Assignments\CSR Forms for ECF\CSR - " & MAXIS_case_number & " on " & file_safe_date & ".docx"
pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CSR Forms for ECF\CSR - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
objDoc.SaveAs pdf_doc_path, 17

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(pdf_doc_path) = TRUE Then
	objDoc.Close wdDoNotSaveChanges
	objWord.Quit

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"
	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "csr-answers-" & MAXIS_case_number & "-info.txt"
	If objFSO.FileExists(local_changelog_path) = True then
		objFSO.DeleteFile(local_changelog_path)
	End If

	Call start_a_blank_case_note
	Call write_variable_in_CASE_NOTE("CSR Form completed via Phone")
	Call write_variable_in_CASE_NOTE("Form information taken verbally per COVID Waiver Allowance.")
	Call write_variable_in_CASE_NOTE("Form information taken on " & csr_form_date)
	Call write_variable_in_CASE_NOTE("CSR for " & MAXIS_footer_month & "/" & MAXIS_footer_year)
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = "Success! The information you have provided for the CSR form has been saved to the Assignments forlder so the CSR Form can be updated and added to ECF. The case can be processed using the information saved in the PDF. Additional notes and information are needed or case processing. This script has NOT updated MAXIS or added CSR processing notes."

	reopen_pdf_doc_msg = MsgBox("The information about the CSR has been saved to a PDF on the LAN to be added to the DHS form and added to ECF." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbYesNo, "Open PDF Doc?")
	If reopen_pdf_doc_msg = vbYes Then
		' Set objAdob = CreateObject("ArcoExch.App")
		' objAdob.Open pdf_doc_path
		' wshshell.Run "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\ArcoRd32.exe", pdf_doc_path
		run_path = chr(34) & pdf_doc_path & chr(34)
		wshshell.Run run_path
		end_msg = end_msg & vbCr & vbCr & "The PDF has been opened for you to view the information that has been saved."
	End If
Else
	end_msg = "Something has gone wrong - the CSR information has NOT been saved correctly to be processed." & vbCr & vbCr & "You can either save the Word Document that has opened as a PDF in the Assignment folder OR Close that document without saving and RERUN the script. Your details have been saved and the script can reopen them and attampt to create the files again. When the script is running, it is best to not interrupt the process."
End If

call script_end_procedure(end_msg)
