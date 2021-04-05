'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/00/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function access_ADDR_panel(access_type, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "ADDR")

        EMReadScreen line_one, 22, 6, 43										'Reading all the information from the panel
        EMReadScreen line_two, 22, 7, 43
        EMReadScreen city_line, 15, 8, 43
        EMReadScreen state_line, 2, 8, 66
        EMReadScreen zip_line, 7, 9, 43
        EMReadScreen county_line, 2, 9, 66
        EMReadScreen verif_line, 2, 9, 74
        EMReadScreen homeless_line, 1, 10, 43
        EMReadScreen reservation_line, 1, 10, 74
        EMReadScreen living_sit_line, 2, 11, 43

        resi_line_one = replace(line_one, "_", "")								'This is all formatting of the information from the panel
        resi_line_two = replace(line_two, "_", "")
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

        EMReadScreen addr_eff_date, 8, 4, 43									'reading the mail information
        EMReadScreen addr_future_date, 8, 4, 66
        EMReadScreen mail_line_one, 22, 13, 43
        EMReadScreen mail_line_two, 22, 14, 43
        EMReadScreen mail_city_line, 15, 15, 43
        EMReadScreen mail_state_line, 2, 16, 43
        EMReadScreen mail_zip_line, 7, 16, 52

        addr_eff_date = replace(addr_eff_date, " ", "/")						'cormatting the mail information
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

        EMReadScreen phone_one, 14, 17, 45										'reading the phone information
        EMReadScreen phone_two, 14, 18, 45
        EMReadScreen phone_three, 14, 19, 45

        EMReadScreen type_one, 1, 17, 67
        EMReadScreen type_two, 1, 18, 67
        EMReadScreen type_three, 1, 19, 67

        phone_one = replace(phone_one, " ) ", "-")								'formatting the phone information
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

function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
'This function will asses the variables provided as the footer month and year to be sure it is correct.
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
'This function records the variables into a txt file so that it can be retrieved by the script if run later.

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	If MAXIS_case_number <> "" Then
		local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	Else
		local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"
	End If
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
			objTextStream.WriteLine "PRE - ATC - " & all_the_clients
			objTextStream.WriteLine "PRE - WHO - " & who_are_we_completing_the_form_with

			objTextStream.WriteLine "EXP - 1 - " & exp_q_1_income_this_month
			objTextStream.WriteLine "EXP - 2 - " & exp_q_2_assets_this_month
			objTextStream.WriteLine "EXP - 3 - RENT - " & exp_q_3_rent_this_month
			objTextStream.WriteLine "EXP - 3 - HEAT - " & exp_pay_heat_checkbox
			objTextStream.WriteLine "EXP - 3 - ACON - " & exp_pay_ac_checkbox
			objTextStream.WriteLine "EXP - 3 - ELEC - " & exp_pay_electricity_checkbox
			objTextStream.WriteLine "EXP - 3 - PHON - " & exp_pay_phone_checkbox
			objTextStream.WriteLine "EXP - 3 - NONE - " & exp_pay_none_checkbox
			objTextStream.WriteLine "EXP - 4 - " & exp_migrant_seasonal_formworker_yn
			objTextStream.WriteLine "EXP - 5 - PREV - " & exp_received_previous_assistance_yn
			objTextStream.WriteLine "EXP - 5 - WHEN - " & exp_previous_assistance_when
			objTextStream.WriteLine "EXP - 5 - WHER - " & exp_previous_assistance_where
			objTextStream.WriteLine "EXP - 5 - WHAT - " & exp_previous_assistance_what
			objTextStream.WriteLine "EXP - 6 - PREG - " & exp_pregnant_yn
			objTextStream.WriteLine "EXP - 6 - WHO? - " & exp_pregnant_who

			objTextStream.WriteLine "ADR - RESI - STR - " & resi_addr_street_full
			objTextStream.WriteLine "ADR - RESI - CIT - " & resi_addr_city
			objTextStream.WriteLine "ADR - RESI - STA - " & resi_addr_state
			objTextStream.WriteLine "ADR - RESI - ZIP - " & resi_addr_zip

			objTextStream.WriteLine "ADR - RESI - RES - " & reservation_yn
			objTextStream.WriteLine "ADR - RESI - NAM - " & reservation_name

			objTextStream.WriteLine "ADR - RESI - HML - " & homeless_yn

			objTextStream.WriteLine "ADR - RESI - LIV - " & living_situation

			objTextStream.WriteLine "ADR - MAIL - STR - " & mail_addr_street_full
			objTextStream.WriteLine "ADR - MAIL - CIT - " & mail_addr_city
			objTextStream.WriteLine "ADR - MAIL - STA - " & mail_addr_state
			objTextStream.WriteLine "ADR - MAIL - ZIP - " & mail_addr_zip

			objTextStream.WriteLine "ADR - PHON - NON - " & phone_one_number
			objTextStream.WriteLine "ADR - PHON - TON - " & phone_pne_type
			objTextStream.WriteLine "ADR - PHON - NTW - " & phone_two_number
			objTextStream.WriteLine "ADR - PHON - TTW - " & phone_two_type
			objTextStream.WriteLine "ADR - PHON - NTH - " & phone_three_number
			objTextStream.WriteLine "ADR - PHON - TTH - " & phone_three_type

			objTextStream.WriteLine "ADR - DATE - " & address_change_date
			objTextStream.WriteLine "ADR - CNTY - " & resi_addr_county

			objTextStream.WriteLine "01A - " & question_1_yn
			objTextStream.WriteLine "01N - " & question_1_notes
			objTextStream.WriteLine "01V - " & question_1_verif_yn
			objTextStream.WriteLine "01D - " & question_1_verif_details

			objTextStream.WriteLine "02A - " & question_2_yn
			objTextStream.WriteLine "02N - " & question_2_notes
			objTextStream.WriteLine "02V - " & question_2_verif_yn
			objTextStream.WriteLine "02D - " & question_2_verif_details

			objTextStream.WriteLine "03A - " & question_3_yn
			objTextStream.WriteLine "03N - " & question_3_notes
			objTextStream.WriteLine "03V - " & question_3_verif_yn
			objTextStream.WriteLine "03D - " & question_3_verif_details

			objTextStream.WriteLine "04A - " & question_4_yn
			objTextStream.WriteLine "04N - " & question_4_notes
			objTextStream.WriteLine "04V - " & question_4_verif_yn
			objTextStream.WriteLine "04D - " & question_4_verif_details

			objTextStream.WriteLine "05A - " & question_5_yn
			objTextStream.WriteLine "05N - " & question_5_notes
			objTextStream.WriteLine "05V - " & question_5_verif_yn
			objTextStream.WriteLine "05D - " & question_5_verif_details

			objTextStream.WriteLine "06A - " & question_6_yn
			objTextStream.WriteLine "06N - " & question_6_notes
			objTextStream.WriteLine "06V - " & question_6_verif_yn
			objTextStream.WriteLine "06D - " & question_6_verif_details

			objTextStream.WriteLine "07A - " & question_7_yn
			objTextStream.WriteLine "07N - " & question_7_notes
			objTextStream.WriteLine "07V - " & question_7_verif_yn
			objTextStream.WriteLine "07D - " & question_7_verif_details

			objTextStream.WriteLine "08A - " & question_8_yn
			objTextStream.WriteLine "08N - " & question_8_notes
			objTextStream.WriteLine "08V - " & question_8_verif_yn
			objTextStream.WriteLine "08D - " & question_8_verif_details

			objTextStream.WriteLine "09A - " & question_9_yn
			objTextStream.WriteLine "09N - " & question_9_notes
			objTextStream.WriteLine "09V - " & question_9_verif_yn
			objTextStream.WriteLine "09D - " & question_9_verif_details

			objTextStream.WriteLine "10A - " & question_10_yn
			objTextStream.WriteLine "10N - " & question_10_notes
			objTextStream.WriteLine "10V - " & question_10_verif_yn
			objTextStream.WriteLine "10D - " & question_10_verif_details
			objTextStream.WriteLine "10G - " & question_10_monthly_earnings

			objTextStream.WriteLine "11A - " & question_11_yn
			objTextStream.WriteLine "11N - " & question_11_notes
			objTextStream.WriteLine "11V - " & question_11_verif_yn
			objTextStream.WriteLine "11D - " & question_11_verif_details

			objTextStream.WriteLine "PWE - " & pwe_selection

			objTextStream.WriteLine "12A - RS - " & question_12_yn
			objTextStream.WriteLine "12A - SS - " & question_12_yn
			objTextStream.WriteLine "12A - VA - " & question_12_yn
			objTextStream.WriteLine "12A - UI - " & question_12_yn
			objTextStream.WriteLine "12A - WC - " & question_12_yn
			objTextStream.WriteLine "12A - RT - " & question_12_yn
			objTextStream.WriteLine "12A - TP - " & question_12_yn
			objTextStream.WriteLine "12A - CS - " & question_12_yn
			objTextStream.WriteLine "12A - OT - " & question_12_yn
			objTextStream.WriteLine "12N - " & question_12_notes
			objTextStream.WriteLine "12V - " & question_12_verif_yn
			objTextStream.WriteLine "12D - " & question_12_verif_details

			objTextStream.WriteLine "13A - " & question_13_yn
			objTextStream.WriteLine "13N - " & question_13_notes
			objTextStream.WriteLine "13V - " & question_13_verif_yn
			objTextStream.WriteLine "13D - " & question_13_verif_details

			objTextStream.WriteLine "14A - RT - " &  question_14_rent_yn
			objTextStream.WriteLine "14A - SB - " &  question_14_subsidy_yn
			objTextStream.WriteLine "14A - MT - " &  question_14_mortgage_yn
			objTextStream.WriteLine "14A - AS - " &  question_14_association_yn
			objTextStream.WriteLine "14A - IN - " &  question_14_insurance_yn
			objTextStream.WriteLine "14A - RM - " &  question_14_room_yn
			objTextStream.WriteLine "14A - TX - " &  question_14_taxes_yn
			objTextStream.WriteLine "14N - " & question_14_notes
			objTextStream.WriteLine "14V - " & question_14_verif_yn
			objTextStream.WriteLine "14D - " & question_14_verif_details

			objTextStream.WriteLine "15A - HA - " & question_15_heat_ac_yn
			objTextStream.WriteLine "15A - EL - " & question_15_electricity_yn
			objTextStream.WriteLine "15A - CF - " & question_15_cooking_fuel_yn
			objTextStream.WriteLine "15A - WS - " & question_15_water_and_sewer_yn
			objTextStream.WriteLine "15A - GR - " & question_15_garbage_yn
			objTextStream.WriteLine "15A - PN - " & question_15_phone_yn
			objTextStream.WriteLine "15A - LP - " & question_15_liheap_yn
			objTextStream.WriteLine "15N - " & question_15_notes
			objTextStream.WriteLine "15V - " & question_15_verif_yn
			objTextStream.WriteLine "15D - " & question_15_verif_details

			objTextStream.WriteLine "16A - " & question_16_yn
			objTextStream.WriteLine "16N - " & question_16_notes
			objTextStream.WriteLine "16V - " & question_16_verif_yn
			objTextStream.WriteLine "16D - " & question_16_verif_details

			objTextStream.WriteLine "17A - " & question_17_yn
			objTextStream.WriteLine "17N - " & question_17_notes
			objTextStream.WriteLine "17V - " & question_17_verif_yn
			objTextStream.WriteLine "17D - " & question_17_verif_details

			objTextStream.WriteLine "18A - " & question_18_yn
			objTextStream.WriteLine "18N - " & question_18_notes
			objTextStream.WriteLine "18V - " & question_18_verif_yn
			objTextStream.WriteLine "18D - " & question_18_verif_details

			objTextStream.WriteLine "19A - " & question_19_yn
			objTextStream.WriteLine "19N - " & question_19_notes
			objTextStream.WriteLine "19V - " & question_19_verif_yn
			objTextStream.WriteLine "19D - " & question_19_verif_details

			objTextStream.WriteLine "20A - CA - " & question_20_cash_yn
			objTextStream.WriteLine "20A - AC - " & question_20_acct_yn
			objTextStream.WriteLine "20A - SE - " & question_20_secu_yn
			objTextStream.WriteLine "20A - CR - " & question_20_cars_yn
			objTextStream.WriteLine "20N - " & question_20_notes
			objTextStream.WriteLine "20V - " & question_20_verif_yn
			objTextStream.WriteLine "20D - " & question_20_verif_details

			objTextStream.WriteLine "21A - " & question_21_yn
			objTextStream.WriteLine "21N - " & question_21_notes
			objTextStream.WriteLine "21V - " & question_21_verif_yn
			objTextStream.WriteLine "21D - " & question_21_verif_details

			objTextStream.WriteLine "22A - " & question_22_yn
			objTextStream.WriteLine "22N - " & question_22_notes
			objTextStream.WriteLine "22V - " & question_22_verif_yn
			objTextStream.WriteLine "22D - " & question_22_verif_details

			objTextStream.WriteLine "23A - " & question_23_yn
			objTextStream.WriteLine "23N - " & question_23_notes
			objTextStream.WriteLine "23V - " & question_23_verif_yn
			objTextStream.WriteLine "23D - " & question_23_verif_details

			objTextStream.WriteLine "24A - RP - " & question_24_rep_payee_yn
			objTextStream.WriteLine "24A - GF - " & question_24_guardian_fees_yn
			objTextStream.WriteLine "24A - SD - " & question_24_special_diet_yn
			objTextStream.WriteLine "24A - HH - " & question_24_high_housing_yn
			objTextStream.WriteLine "24N - " & question_24_notes
			objTextStream.WriteLine "24V - " & question_24_verif_yn
			objTextStream.WriteLine "24D - " & question_24_verif_details

			objTextStream.WriteLine "QQ1A - " & qual_question_one
			objTextStream.WriteLine "QQ1M - " & qual_memb_one
			objTextStream.WriteLine "QQ2A - " & qual_question_two
			objTextStream.WriteLine "QQ2M - " & qual_memb_two
			objTextStream.WriteLine "QQ3A - " & qual_question_three
			objTextStream.WriteLine "QQ3M - " & qual_memb_there
			objTextStream.WriteLine "QQ4A - " & qual_question_four
			objTextStream.WriteLine "QQ4M - " & qual_memb_four
			objTextStream.WriteLine "QQ5A - " & qual_question_five
			objTextStream.WriteLine "QQ5M - " & qual_memb_five

			For known_membs = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
				objTextStream.WriteLine "ARR - ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_last_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_first_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_other_names, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_dob, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_gender, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_former_state, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_citizen, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_written_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_notes, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
			Next

			for this_jobs = 0 to UBOUND(JOBS_ARRAY, 2)
				objTextStream.WriteLine "ARR - JOBS_ARRAY - " & JOBS_ARRAY(jobs_employee_name, this_jobs)&"~"&JOBS_ARRAY(jobs_hourly_wage, this_jobs)&"~"&JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)&"~"&JOBS_ARRAY(jobs_employer_name, this_jobs)&"~"&JOBS_ARRAY(jobs_notes, this_jobs)
			Next

			'Close the object so it can be opened again shortly
			objTextStream.Close

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

function restore_your_work(vars_filled)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run
'
	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now determines name of file
	If MAXIS_case_number <> "" Then local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	If no_case_number_checkbox = checked Then local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"

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
				saved_caf_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE

				array_counters = 0
				known_membs = 0
				known_jobs = 0
				For Each text_line in saved_caf_details
					' MsgBox "~" & left(text_line, 9) & "~"
					' MsgBox text_line
					If left(text_line, 9) = "PRE - WHO" Then who_are_we_completing_the_form_with = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - ATC" Then all_the_clients = Mid(text_line, 13)
					If left(text_line, 7) = "EXP - 1" Then exp_q_1_income_this_month = Mid(text_line, 11)
					If left(text_line, 7) = "EXP - 2" Then exp_q_2_assets_this_month = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 3 - RENT" Then exp_q_3_rent_this_month = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - HEAT" Then exp_pay_heat_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - ACON" Then exp_pay_ac_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - ELEC" Then exp_pay_electricity_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - PHON" Then exp_pay_phone_checkbox = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - NONE" Then exp_pay_none_checkbox = Mid(text_line, 18)
					If left(text_line, 7) = "EXP - 4" Then exp_migrant_seasonal_formworker_yn = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 5 - PREV" Then exp_received_previous_assistance_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHEN" Then exp_previous_assistance_when = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHER" Then exp_previous_assistance_where = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHAT" Then exp_previous_assistance_what = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - PREG" Then exp_pregnant_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - WHO?" Then exp_pregnant_who = Mid(text_line, 18)
					If left(text_line, 3) = "ADR" Then
						' MsgBox "~" & mid(text_line, 7, 10) & "~"
						If mid(text_line, 7, 10) = "RESI - STR" Then resi_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - CIT" Then resi_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - STA" Then resi_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - ZIP" Then resi_addr_zip = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - RES" Then reservation_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - NAM" Then reservation_name = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - HML" Then homeless_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - LIV" Then living_situation = MID(text_line, 20)

						If mid(text_line, 7, 10) = "MAIL - STR" Then mail_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - CIT" Then mail_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - STA" Then mail_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - ZIP" Then mail_addr_zip = MID(text_line, 20)

						If mid(text_line, 7, 10) = "PHON - NON" Then phone_one_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TON" Then phone_one_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTW" Then phone_two_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTW" Then phone_two_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTH" Then phone_three_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTH" Then phone_three_type = MID(text_line, 20)

						If mid(text_line, 7, 4) = "DATE" Then address_change_date = MID(text_line, 13)
						If mid(text_line, 7, 4) = "CNTY" Then resi_addr_county = MID(text_line, 13)

					End If
					' If left(text_line, 3) = "" Then  = Mid(text_line, 7)
					If left(text_line, 3) = "01A" Then question_1_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01N" Then question_1_notes = Mid(text_line, 7)
					If left(text_line, 3) = "01V" Then question_1_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01D" Then question_1_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "02A" Then question_2_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02N" Then question_2_notes = Mid(text_line, 7)
					If left(text_line, 3) = "02V" Then question_2_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02D" Then question_2_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "03A" Then question_3_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03N" Then question_3_notes = Mid(text_line, 7)
					If left(text_line, 3) = "03V" Then question_3_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03D" Then question_3_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "04A" Then question_4_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04N" Then question_4_notes = Mid(text_line, 7)
					If left(text_line, 3) = "04V" Then question_4_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04D" Then question_4_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "05A" Then question_5_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05N" Then question_5_notes = Mid(text_line, 7)
					If left(text_line, 3) = "05V" Then question_5_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05D" Then question_5_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "06A" Then question_6_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06N" Then question_6_notes = Mid(text_line, 7)
					If left(text_line, 3) = "06V" Then question_6_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06D" Then question_6_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "07A" Then question_7_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07N" Then question_7_notes = Mid(text_line, 7)
					If left(text_line, 3) = "07V" Then question_7_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07D" Then question_7_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "08A" Then question_8_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08N" Then question_8_notes = Mid(text_line, 7)
					If left(text_line, 3) = "08V" Then question_8_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08D" Then question_8_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "09A" Then question_9_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09N" Then question_9_notes = Mid(text_line, 7)
					If left(text_line, 3) = "09V" Then question_9_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09D" Then question_9_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "10A" Then question_10_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10N" Then question_10_notes = Mid(text_line, 7)
					If left(text_line, 3) = "10V" Then question_10_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10D" Then question_10_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "10G" Then question_10_monthly_earnings = Mid(text_line, 7)

					If left(text_line, 3) = "11A" Then question_11_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11N" Then question_11_notes = Mid(text_line, 7)
					If left(text_line, 3) = "11V" Then question_11_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11D" Then question_11_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "PWE" Then pwe_selection = Mid(text_line, 7)

					If left(text_line, 8) = "12A - RS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - SS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - VA" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - UI" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - WC" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - RT" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - TP" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - CS" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12A - OT" Then question_12_yn = Mid(text_line, 12)
					If left(text_line, 3) = "12N" Then question_12_notes = Mid(text_line, 7)
					If left(text_line, 3) = "12V" Then question_12_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "12D" Then question_12_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "13A" Then question_13_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13N" Then question_13_notes = Mid(text_line, 7)
					If left(text_line, 3) = "13V" Then question_13_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13D" Then question_13_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "14A - RT" Then  question_14_rent_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - SB" Then  question_14_subsidy_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - MT" Then  question_14_mortgage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - AS" Then  question_14_association_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - IN" Then  question_14_insurance_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - RM" Then  question_14_room_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - TX" Then  question_14_taxes_yn = Mid(text_line, 12)
					If left(text_line, 3) = "14N" Then question_14_notes = Mid(text_line, 7)
					If left(text_line, 3) = "14V" Then question_14_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "14D" Then question_14_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "15A - HA" Then question_15_heat_ac_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - EL" Then question_15_electricity_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - CF" Then question_15_cooking_fuel_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - WS" Then question_15_water_and_sewer_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - GR" Then question_15_garbage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - PN" Then question_15_phone_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - LP" Then question_15_liheap_yn = Mid(text_line, 12)
					If left(text_line, 3) = "15N" Then question_15_notes = Mid(text_line, 7)
					If left(text_line, 3) = "15V" Then question_15_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "15D" Then question_15_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "16A" Then question_16_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16N" Then question_16_notes = Mid(text_line, 7)
					If left(text_line, 3) = "16V" Then question_16_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16D" Then question_16_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "17A" Then question_17_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17N" Then question_17_notes = Mid(text_line, 7)
					If left(text_line, 3) = "17V" Then question_17_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17D" Then question_17_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "18A" Then question_18_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18N" Then question_18_notes = Mid(text_line, 7)
					If left(text_line, 3) = "18V" Then question_18_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18D" Then question_18_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "19A" Then question_19_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19N" Then question_19_notes = Mid(text_line, 7)
					If left(text_line, 3) = "19V" Then question_19_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19D" Then question_19_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "20A - CA" Then question_20_cash_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - AC" Then question_20_acct_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - SE" Then question_20_secu_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - CR" Then question_20_cars_yn = Mid(text_line, 12)
					If left(text_line, 3) = "20N" Then question_20_notes = Mid(text_line, 7)
					If left(text_line, 3) = "20V" Then question_20_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "20D" Then question_20_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "21A" Then question_21_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21N" Then question_21_notes = Mid(text_line, 7)
					If left(text_line, 3) = "21V" Then question_21_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21D" Then question_21_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "22A" Then question_22_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22N" Then question_22_notes = Mid(text_line, 7)
					If left(text_line, 3) = "22V" Then question_22_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22D" Then question_22_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "23A" Then question_23_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23N" Then question_23_notes = Mid(text_line, 7)
					If left(text_line, 3) = "23V" Then question_23_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23D" Then question_23_verif_details = Mid(text_line, 7)

					If left(text_line, 8) = "24A - RP" Then question_24_rep_payee_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - GF" Then question_24_guardian_fees_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - SD" Then question_24_special_diet_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - HH" Then question_24_high_housing_yn = Mid(text_line, 12)
					If left(text_line, 3) = "24N" Then question_24_notes = Mid(text_line, 7)
					If left(text_line, 3) = "24V" Then question_24_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "24D" Then question_24_verif_details = Mid(text_line, 7)

					If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ1M" Then qual_memb_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2A" Then qual_question_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2M" Then qual_memb_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3A" Then qual_question_three = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3M" Then qual_memb_there = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4A" Then qual_question_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4M" Then qual_memb_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5A" Then qual_question_five = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5M" Then qual_memb_five = Mid(text_line, 8)

					' If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)

					If left(text_line, 3) = "ARR" Then
						If MID(text_line, 7, 17) = "ALL_CLIENTS_ARRAY" Then
							array_info = Mid(text_line, 27)
							array_info = split(array_info, "~")
							ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, known_membs)
							ALL_CLIENTS_ARRAY(memb_last_name, known_membs) 				= array_info(0)
							ALL_CLIENTS_ARRAY(memb_first_name, known_membs) 			= array_info(1)
							ALL_CLIENTS_ARRAY(memb_mid_name, known_membs) 				= array_info(2)
							ALL_CLIENTS_ARRAY(memb_other_names, known_membs) 			= array_info(3)
							ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs) 				= array_info(4)
							ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs) 			= array_info(5)
							ALL_CLIENTS_ARRAY(memb_dob, known_membs) 					= array_info(6)
							ALL_CLIENTS_ARRAY(memb_gender, known_membs) 				= array_info(7)
							ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs) 			= array_info(8)
							ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs) 		= array_info(9)
							ALL_CLIENTS_ARRAY(memi_last_grade, known_membs) 			= array_info(10)
							ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs) 			= array_info(11)
							ALL_CLIENTS_ARRAY(memi_former_state, known_membs) 			= array_info(12)
							ALL_CLIENTS_ARRAY(memi_citizen, known_membs) 				= array_info(13)
							ALL_CLIENTS_ARRAY(memb_interpreter, known_membs) 			= array_info(14)
							ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs) 		= array_info(15)
							ALL_CLIENTS_ARRAY(memb_written_language, known_membs) 		= array_info(16)
							ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs) 				= array_info(17)
							ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs) 		= array_info(18)
							ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs) 		= array_info(19)
							ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs) 		= array_info(20)
							ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs) 		= array_info(21)
							ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs) 		= array_info(22)
							ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs) 			= array_info(23)
							ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs) 			= array_info(24)
							ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs) 			= array_info(25)
							ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs) 			= array_info(26)
							ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs) 	= array_info(27)
							ALL_CLIENTS_ARRAY(clt_imig_status, known_membs) 			= array_info(28)
							ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs) 				= array_info(29)
							ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs) 				= array_info(30)
							ALL_CLIENTS_ARRAY(clt_verif_details, known_membs) 			= array_info(31)
							ALL_CLIENTS_ARRAY(memb_notes, known_membs) 					= array_info(32)
							ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs) 				= array_info(33)
							known_membs = known_membs + 1
						End If

						If MID(text_line, 7, 10) = "JOBS_ARRAY" Then
							array_info = Mid(text_line, 20)
							array_info = split(array_info, "~")
							ReDim Preserve JOBS_ARRAY(jobs_notes, known_jobs)
							JOBS_ARRAY(jobs_employee_name, known_jobs) 			= array_info(0)
							JOBS_ARRAY(jobs_hourly_wage, known_jobs) 			= array_info(1)
							JOBS_ARRAY(jobs_gross_monthly_earnings, known_jobs) = array_info(2)
							JOBS_ARRAY(jobs_employer_name, known_jobs) 			= array_info(3)
							JOBS_ARRAY(jobs_notes, known_jobs) 					= array_info(4)
							known_jobs = known_jobs + 1
						End If
					End If
				Next
			End If
		End If
	End With
end function

'THESE FUNCTIONS ARE ALL THE INDIVIDUAL DIALOGS WITHIN THE MAIN DIALOG LOOP
function dlg_page_one_pers_and_exp()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 416, 240, "CAF Person and Expedited"
			  ComboBox 205, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_form_with, who_are_we_completing_the_form_with
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
			    CancelButton 360, 220, 50, 15
			  Text 70, 15, 130, 10, "Who are you completing the form with?"
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
	If resi_addr_county = "" Then resi_addr_county = "27 Hennepin"
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
				  DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", reservation_yn
				  EditBox 245, 90, 130, 15, reservation_name
				  DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", homeless_yn
				  DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
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
			Call validate_phone_number(err_msg, "*", phone_one_number, TRUE)
			Call validate_phone_number(err_msg, "*", phone_two_number, TRUE)
			Call validate_phone_number(err_msg, "*", phone_three_number, TRUE)

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

			If shown_known_pers_detail = TRUE Then
				BeginDialog Dialog1, 0, 0, 550, 385, "Household Member Information"
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
				  DropListBox 15, 230, 80, 45, "Yes"+chr(9)+"No", ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)
				  EditBox 100, 230, 205, 15, ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)
				  DropListBox 310, 230, 55, 45, "No"+chr(9)+"Yes", ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)
				  DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
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
						If the_memb = known_membs Then
							Text 498, y_pos + 1, 45, 10, "Person " & (the_memb + 1)
						Else
							PushButton 490, y_pos, 45, 10, "Person " & (the_memb + 1), ALL_CLIENTS_ARRAY(clt_nav_btn, the_memb)
						End If
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

				BeginDialog Dialog1, 0, 0, 550, 385, "Household Member Information"
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
				  DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)
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
			If ButtonPressed = add_person_btn Then
				last_clt = UBound(ALL_CLIENTS_ARRAY, 2)
				new_clt = last_clt + 1
				ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, new_clt)
				known_membs = new_clt
			End If
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
			BeginDialog Dialog1, 0, 0, 550, 385, "Tell Us About Your Household"
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

function dlg_page_four_income()
	go_back = FALSE
	Do
		Do
			err_msg = ""

			btn_placeholder = 4000
			dlg_len = 350
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				JOBS_ARRAY(jobs_edit_btn, each_job) = btn_placeholder
				btn_placeholder = btn_placeholder + 1
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then dlg_len = dlg_len + 10
			next

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 650, dlg_len, "What kinds of income do you have?"
			  DropListBox 10, 10, 60, 45, question_answers, question_8_yn
			  Text 80, 10, 290, 10, "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			  Text 540, 10, 105, 10, "Q8 - Verification - " & question_8_verif_yn
			  ButtonGroup ButtonPressed
			    PushButton 560, 20, 75, 10, "ADD VERIFICATION", add_verif_8_btn
			  DropListBox 10, 25, 60, 45, question_answers, question_8a_yn
			  Text 90, 25, 350, 10, "a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			  Text 95, 40, 25, 10, "Notes:"
			  EditBox 120, 35, 390, 15, question_8_notes
			  DropListBox 10, 55, 60, 45, question_answers, question_9_yn
			  Text 80, 55, 350, 10, "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			  ButtonGroup ButtonPressed
			    PushButton 430, 55, 55, 10, "ADD JOB", add_job_btn
			  Text 540, 55, 105, 10, "Q9 - Verification - " & question_9_verif_yn
			  ButtonGroup ButtonPressed
			    PushButton 560, 65, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			  y_pos = 65
			  ' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			  for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				  ' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				  If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then

					  Text 95, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					  ButtonGroup ButtonPressed
					    PushButton 495, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					  y_pos = y_pos + 10
				  End If
			  next
			  y_pos = y_pos + 10
			  Text 95, y_pos, 25, 10, "Notes:"
			  EditBox 120, y_pos - 5, 390, 15, question_9_notes
			  y_pos = y_pos + 15
			  DropListBox 10, y_pos, 60, 45, question_answers, question_10_yn
			  Text 80, y_pos, 430, 10, "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			  Text 540, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_10_btn
			  Text 95, y_pos, 85, 10, "Gross Monthly Earnings:"
			  Text 185, y_pos, 25, 10, "Notes:"
			  y_pos = y_pos + 10
			  EditBox 95, y_pos, 80, 15, question_10_monthly_earnings
			  EditBox 185, y_pos, 325, 15, question_10_notes
			  y_pos = y_pos + 20
			  DropListBox 10, y_pos, 60, 45, question_answers, question_11_yn
			  Text 80, y_pos, 255, 10, "11. Do you expect any changes in income, expenses or work hours?"
			  Text 540, y_pos, 105, 10, "Q11 - Verification - " & question_11_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_11_btn
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_11_notes
			  y_pos = y_pos + 25
			  Text 80, y_pos, 75, 10, "Pricipal Wage Earner"
			  DropListBox 155, y_pos - 5, 175, 45, all_the_clients, pwe_selection
			  y_pos = y_pos + 10
			  Text 80, y_pos + 5, 370, 10, "12. Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			  Text 540, y_pos, 105, 10, "Q12 - Verification - " & question_12_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
			    PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_12_btn
			  y_pos = y_pos + 10
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_rsdi_yn
			  Text 150, y_pos + 5, 70, 10, "RSDI                      $"
			  EditBox 220, y_pos, 35, 15, question_12_rsdi_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ssi_yn
			  Text 375, y_pos + 5, 85, 10, "SSI                                 $"
			  EditBox 460, y_pos, 35, 15, question_12_ssi_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_va_yn
			  Text 150, y_pos + 5, 70, 10, "VA                          $"
			  EditBox 220, y_pos, 35, 15, question_12_va_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ui_yn
			  Text 375, y_pos + 5, 85, 10, "Unemployment Ins          $"
			  EditBox 460, y_pos, 35, 15, question_12_ui_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_wc_yn
			  Text 150, y_pos + 5, 70, 10, "Workers Comp       $"
			  EditBox 220, y_pos, 35, 15, question_12_wc_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_ret_yn
			  Text 375, y_pos + 5, 85, 10, "Retirement Ben.              $"
			  EditBox 460, y_pos, 35, 15, question_12_ret_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_trib_yn
			  Text 150, y_pos + 5, 70, 10, "Tribal Payments      $"
			  EditBox 220, y_pos, 35, 15, question_12_trib_amt
			  DropListBox 305, y_pos, 60, 45, question_answers, question_12_cs_yn
			  Text 375, y_pos + 5, 85, 10, "Child/Spousal Support    $"
			  EditBox 460, y_pos, 35, 15, question_12_cs_amt
			  y_pos = y_pos + 15
			  DropListBox 80, y_pos, 60, 45, question_answers, question_12_other_yn
			  Text 150, y_pos + 5, 110, 10, "Other unearned income          $"
			  EditBox 250, y_pos, 35, 15, question_12_other_amt
			  y_pos = y_pos + 20
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_12_notes
			  y_pos = y_pos + 25
			  DropListBox 10, y_pos, 60, 45, question_answers, question_13_yn
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 80, y_pos, 400, 10, "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			  Text 540, y_pos, 105, 10, "Q13 - Verification - " & question_13_verif_yn
			  y_pos = y_pos + 10
			  ButtonGroup ButtonPressed
				PushButton 560, y_pos, 75, 10, "ADD VERIFICATION", add_verif_13_btn
			  Text 95, y_pos + 5, 25, 10, "Notes:"
			  EditBox 120, y_pos, 390, 15, question_13_notes
			  y_pos = y_pos + 20
			  ButtonGroup ButtonPressed
			    PushButton 540, y_pos, 50, 15, "Next", next_btn
			    PushButton 485, y_pos + 5, 50, 10, "Back", back_btn
			    CancelButton 595, y_pos, 50, 15
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_8_btn Then Call verif_details_dlg(8)
			If ButtonPressed = add_verif_9_btn Then Call verif_details_dlg(9)
			If ButtonPressed = add_verif_10_btn Then Call verif_details_dlg(10)
			If ButtonPressed = add_verif_11_btn Then Call verif_details_dlg(11)
			If ButtonPressed = add_verif_12_btn Then Call verif_details_dlg(12)
			If ButtonPressed = add_verif_13_btn Then Call verif_details_dlg(13)

			If ButtonPressed = add_job_btn Then
				another_job = ""
				count = 0
				for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
					count = count + 1
					If JOBS_ARRAY(jobs_employer_name, each_job) = "" AND JOBS_ARRAY(jobs_employee_name, each_job) = "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) = "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) = "" Then
						another_job = each_job
					End If
				Next
				If another_job = "" Then
					another_job = count
					ReDim Preserve JOBS_ARRAY(jobs_notes, another_job)
				End If
				Call jobs_details_dlg(another_job)
			End If

			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				If ButtonPressed = JOBS_ARRAY(jobs_edit_btn, each_job) Then
					Call jobs_details_dlg(each_job)
				End If
			next

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_3_hhinfo_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_4_income_dlg = FALSE
		caf_pg_4_income_dlg_cleared = TRUE
	End If

end function

function dlg_page_five_expenses()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""

			BeginDialog Dialog1, 0, 0, 550, 385, "What kinds of expenses do you have?"
			  DropListBox 95, 20, 60, 45, question_answers, question_14_rent_yn
			  DropListBox 300, 20, 60, 45, question_answers, question_14_subsidy_yn
			  DropListBox 95, 35, 60, 45, question_answers, question_14_mortgage_yn
			  DropListBox 300, 35, 60, 45, question_answers, question_14_association_yn
			  DropListBox 95, 50, 60, 45, question_answers, question_14_insurance_yn
			  DropListBox 300, 50, 60, 45, question_answers, question_14_room_yn
			  DropListBox 95, 65, 60, 45, question_answers, question_14_taxes_yn
			  EditBox 135, 85, 390, 15, question_14_notes
			  DropListBox 95, 120, 60, 45, question_answers, question_15_heat_ac_yn
			  DropListBox 265, 120, 60, 45, question_answers, question_15_electricity_yn
			  DropListBox 415, 120, 60, 45, question_answers, question_15_cooking_fuel_yn
			  DropListBox 95, 135, 60, 45, question_answers, question_15_water_and_sewer_yn
			  DropListBox 265, 135, 60, 45, question_answers, question_15_garbage_yn
			  DropListBox 415, 135, 60, 45, question_answers, question_15_phone_yn
			  DropListBox 95, 150, 60, 45, question_answers, question_15_liheap_yn
			  EditBox 120, 165, 390, 15, question_15_notes
			  DropListBox 10, 190, 60, 45, question_answers, question_16_yn
			  EditBox 120, 210, 390, 15, question_16_notes
			  DropListBox 10, 235, 60, 45, question_answers, question_17_yn
			  EditBox 120, 255, 390, 15, question_17_notes
			  DropListBox 10, 280, 60, 45, question_answers, question_18_yn
			  EditBox 120, 300, 390, 15, question_18_notes
			  DropListBox 10, 325, 60, 45, question_answers, question_19_yn
			  EditBox 120, 335, 390, 15, question_19_notes
			  ButtonGroup ButtonPressed
			    PushButton 560, 355, 50, 15, "Next", next_btn
			    PushButton 505, 360, 50, 10, "Back", back_btn
			    CancelButton 615, 355, 50, 15
			    PushButton 580, 20, 75, 10, "ADD VERIFICATION", add_verif_14_btn
			    PushButton 580, 120, 75, 10, "ADD VERIFICATION", add_verif_15_btn
			    PushButton 580, 200, 75, 10, "ADD VERIFICATION", add_verif_16_btn
			    PushButton 580, 245, 75, 10, "ADD VERIFICATION", add_verif_17_btn
			    PushButton 580, 290, 75, 10, "ADD VERIFICATION", add_verif_18_btn
			    PushButton 580, 335, 75, 10, "ADD VERIFICATION", add_verif_19_btn
			  Text 80, 10, 220, 10, "14. Does your household have the following housing expenses?"
			  Text 165, 25, 70, 10, "Rent"
			  Text 370, 25, 100, 10, "Rent or Section 8 Subsidy"
			  Text 165, 40, 125, 10, "Mortgage/contract for deed payment"
			  Text 370, 40, 70, 10, "Association fees"
			  Text 165, 55, 85, 10, "Homeowner's insurance"
			  Text 370, 55, 70, 10, "Room and/or board"
			  Text 165, 70, 100, 10, "Real estate taxes"
			  Text 110, 90, 25, 10, "Notes:"
			  Text 560, 10, 105, 10, "Q14 - Verification - " & question_14_verif_yn
			  Text 80, 110, 290, 10, "15. Does your household have the following utility expenses any time during the year? "
			  Text 165, 125, 85, 10, "Heating/air conditioning"
			  Text 335, 125, 70, 10, "Electricity"
			  Text 485, 125, 70, 10, "Cooking fuel"
			  Text 165, 140, 75, 10, "Water and sewer"
			  Text 335, 140, 60, 10, "Garbage removal"
			  Text 485, 140, 70, 10, "Phone/cell phone"
			  Text 165, 155, 375, 10, "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"
			  Text 95, 170, 25, 10, "Notes:"
			  Text 560, 110, 105, 10, "Q15 - Verification - " & question_15_verif_yn
			  Text 80, 190, 345, 10, "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working,"
			  Text 95, 200, 125, 10, "looking for work or going to school?"
			  Text 95, 215, 25, 10, "Notes:"
			  Text 560, 190, 105, 10, "Q16 - Verification - " & question_16_verif_yn
			  Text 80, 235, 380, 10, "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working,"
			  Text 95, 245, 125, 10, "looking for work or going to school?"
			  Text 95, 260, 25, 10, "Notes:"
			  Text 560, 235, 105, 10, "Q17 - Verification - " & question_17_verif_yn
			  Text 80, 280, 430, 10, "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support"
			  Text 95, 290, 215, 10, "or contribute to a tax dependent who does not live in your home?"
			  Text 95, 305, 25, 10, "Notes:"
			  Text 0, 0, 0, 0, ""
			  Text 0, 0, 0, 0, ""
			  Text 560, 280, 105, 10, "Q18 - Verification - " & question_18_verif_yn
			  Text 80, 325, 255, 10, "19. For SNAP only: Does anyone in the household have medical expenses? "
			  Text 95, 340, 25, 10, "Notes:"
			  Text 560, 325, 105, 10, "Q19 - Verification - " & question_19_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_14_btn Then Call verif_details_dlg(14)
			If ButtonPressed = add_verif_15_btn Then Call verif_details_dlg(15)
			If ButtonPressed = add_verif_16_btn Then Call verif_details_dlg(16)
			If ButtonPressed = add_verif_17_btn Then Call verif_details_dlg(17)
			If ButtonPressed = add_verif_18_btn Then Call verif_details_dlg(18)
			If ButtonPressed = add_verif_19_btn Then Call verif_details_dlg(19)


			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_4_income_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_5_expenses_dlg = FALSE
		caf_pg_5_expenses_dlg_cleared = TRUE
	End If

end function

function dlg_page_six_assets_and_other()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "What do you own? Other Information"
			  DropListBox 80, 25, 60, 45, question_answers, question_20_cash_yn
			  DropListBox 285, 25, 60, 45, question_answers, question_20_acct_yn
			  DropListBox 80, 40, 60, 45, question_answers, question_20_secu_yn
			  DropListBox 285, 40, 60, 45, question_answers, question_20_cars_yn
			  EditBox 120, 60, 390, 15, question_20_notes
			  DropListBox 10, 85, 60, 45, question_answers, question_21_yn
			  EditBox 120, 95, 390, 15, question_21_notes
			  DropListBox 10, 120, 60, 45, question_answers, question_22_yn
			  EditBox 120, 130, 390, 15, question_22_notes
			  DropListBox 10, 155, 60, 45, question_answers, question_23_yn
			  EditBox 120, 165, 390, 15, question_23_notes
			  DropListBox 80, 205, 60, 45, question_answers, question_24_rep_payee_yn
			  DropListBox 285, 205, 60, 45, question_answers, question_24_guardian_fees_yn
			  DropListBox 80, 220, 60, 45, question_answers, question_24_special_diet_yn
			  DropListBox 285, 220, 60, 45, question_answers, question_24_high_housing_yn
			  EditBox 120, 240, 390, 15, question_24_notes
			  ButtonGroup ButtonPressed
			    PushButton 540, 240, 50, 15, "Next", next_btn
			    PushButton 540, 230, 50, 10, "Back", back_btn
			    CancelButton 595, 240, 50, 15
			    PushButton 560, 20, 75, 10, "ADD VERIFICATION", add_verif_20_btn
			    PushButton 560, 95, 75, 10, "ADD VERIFICATION", add_verif_21_btn
			    PushButton 560, 130, 75, 10, "ADD VERIFICATION", add_verif_22_btn
			    PushButton 560, 165, 75, 10, "ADD VERIFICATION", add_verif_23_btn
			    PushButton 560, 200, 75, 10, "ADD VERIFICATION", add_verif_24_btn
			  Text 80, 10, 280, 10, "20. Does anyone in the household own, or is anyone buying, any of the following?"
			  Text 150, 30, 70, 10, "Cash"
			  Text 355, 30, 175, 10, "Bank accounts (savings, checking, debit card, etc.)"
			  Text 150, 45, 125, 10, "Stocks, bonds, annuities, 401k, etc."
			  Text 355, 45, 180, 10, "Vehicles (cars, trucks, motorcycles, campers, trailers)"
			  Text 95, 65, 25, 10, "Notes:"
			  Text 540, 10, 105, 10, "Q20 - Verification - " & question_20_verif_yn
			  Text 80, 85, 420, 10, "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? "
			  Text 95, 100, 25, 10, "Notes:"
			  Text 540, 85, 105, 10, "Q21 - Verification - " & question_21_verif_yn
			  Text 80, 120, 305, 10, "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			  Text 95, 135, 25, 10, "Notes:"
			  Text 540, 120, 105, 10, "Q22 - Verification - " & question_22_verif_yn
			  Text 80, 155, 250, 10, "23. For children under the age of 19, are both parents living in the home?"
			  Text 95, 170, 25, 10, "Notes:"
			  Text 540, 155, 105, 10, "Q23 - Verification - " & question_23_verif_yn
			  Text 80, 190, 325, 10, "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			  Text 150, 210, 95, 10, "Representative Payee fees"
			  Text 355, 210, 105, 10, "Guardian Conservator fees"
			  Text 150, 225, 125, 10, "Physician-perscribed special diet"
			  Text 355, 225, 105, 10, "High housing costs"
			  Text 95, 245, 25, 10, "Notes:"
			  Text 540, 190, 105, 10, "Q24 - Verification - " & question_24_verif_yn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			'ADD ERROR HANDLING HERE

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = add_verif_20_btn Then Call verif_details_dlg(14)
			If ButtonPressed = add_verif_21_btn Then Call verif_details_dlg(15)
			If ButtonPressed = add_verif_22_btn Then Call verif_details_dlg(16)
			If ButtonPressed = add_verif_23_btn Then Call verif_details_dlg(17)
			If ButtonPressed = add_verif_24_btn Then Call verif_details_dlg(18)


			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_5_expenses_dlg = TRUE
			End If

		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_pg_6_other_dlg = FALSE
		caf_pg_6_other_dlg_cleared = TRUE
	End If
end function


function dlg_qualifying_questions()
	go_back = FALSE
	Do
		Do
			err_msg = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "CAF Qualifying Questions"
			  DropListBox 220, 40, 30, 45, "No"+chr(9)+"Yes", qual_question_one
			  ComboBox 340, 40, 105, 45, all_the_clients, qual_memb_one
			  DropListBox 220, 80, 30, 45, "No"+chr(9)+"Yes", qual_question_two
			  ComboBox 340, 80, 105, 45, all_the_clients, qual_memb_two
			  DropListBox 220, 110, 30, 45, "No"+chr(9)+"Yes", qual_question_three
			  ComboBox 340, 110, 105, 45, all_the_clients, qual_memb_there
			  DropListBox 220, 140, 30, 45, "No"+chr(9)+"Yes", qual_question_four
			  ComboBox 340, 140, 105, 45, all_the_clients, qual_memb_four
			  DropListBox 220, 160, 30, 45, "No"+chr(9)+"Yes", qual_question_five
			  ComboBox 340, 160, 105, 45, all_the_clients, qual_memb_five
			  ButtonGroup ButtonPressed
			    CancelButton 395, 185, 50, 15
			    PushButton 340, 185, 50, 15, "Next", next_btn
			    PushButton 285, 190, 50, 10, "Back", back_btn
			  Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
			  Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			  Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			  Text 10, 110, 195, 30, "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			  Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
			  Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			  Text 260, 40, 70, 10, "Household Member:"
			  Text 260, 80, 70, 10, "Household Member:"
			  Text 260, 110, 70, 10, "Household Member:"
			  Text 260, 140, 70, 10, "Household Member:"
			  Text 260, 160, 70, 10, "Household Member:"
			EndDialog

			dialog Dialog1

			cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = next_btn

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = next_btn
				err_msg = ""
				show_caf_pg_5_expenses_dlg = TRUE
			End If
		Loop until ButtonPressed = next_btn
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_qual_questions_dlg = FALSE
		caf_caf_qual_questions_dlg_cleared = TRUE
	End If

end function


function dlg_signature()
	go_back = FALSE
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "Form dates and signatures"
			  EditBox 135, 50, 60, 15, caf_form_date
			  DropListBox 135, 70, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_yn
			  ButtonGroup ButtonPressed
			    PushButton 35, 90, 105, 15, "Complete CAF Form Detail", complete_caf_questions
			    PushButton 5, 95, 25, 10, "BACK", back_btn
			    PushButton 10, 35, 145, 10, "Open RIGHTS AND RESPONSIBLITIES ", open_r_and_r_button
			    CancelButton 145, 90, 50, 15
			  Text 10, 10, 160, 20, "Confirm the client is signing this form and attesting to the information provided verbally."
			  Text 70, 55, 55, 10, "CAF Form Date:"
			  Text 10, 75, 120, 10, "Cient signature accepted verbally?"
			EndDialog

			dialog Dialog1

			cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = complete_caf_questions
			If ButtonPressed = open_r_and_r_button Then open_URL_in_browser("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG")

			If IsDate(caf_form_date) = FALSE Then
				err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was received."
			Else
				If DateDiff("d", date, caf_form_date) > 0 Then err_msg = err_msg & vbNewLine & "* The date of the CAF form is listed as a future date, a form cannot be listed as received inthe future, please review the form date."
			End If
			If client_signed_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the client has signed the form correctly by selecting 'yes' or 'no'."

			If ButtonPressed = back_btn Then
				go_back = TRUE
				ButtonPressed = complete_caf_questions
				err_msg = ""
				show_caf_qual_questions_dlg = TRUE
			End If
		Loop until ButtonPressed = complete_caf_questions
		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
	Loop until err_msg = ""

	If go_back = FALSE Then
		show_caf_sig_dlg = FALSE
		caf_sig_dlg_cleared = TRUE
	End If
end function

function verif_details_dlg(question_number)
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


	BeginDialog Dialog1, 0, 0, 550, 385, "Add Verification"
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

function jobs_details_dlg(this_jobs)
	Do
		pick_a_client = replace(all_the_clients, "Select or Type", "Select One...")
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 321, 130, "Add Job"
		  DropListBox 10, 35, 135, 45, pick_a_client+chr(9)+"", JOBS_ARRAY(jobs_employee_name, this_jobs)
		  EditBox 150, 35, 60, 15, JOBS_ARRAY(jobs_hourly_wage, this_jobs)
		  EditBox 215, 35, 100, 15, JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)
		  EditBox 10, 65, 305, 15, JOBS_ARRAY(jobs_employer_name, this_jobs)
		  EditBox 35, 90, 280, 15, JOBS_ARRAY(jobs_notes, this_jobs)
		  ButtonGroup ButtonPressed
		    PushButton 265, 110, 50, 15, "Return", return_btn
		    PushButton 265, 10, 50, 10, "CLEAR", clear_job_btn
		  Text 10, 10, 100, 10, "Enter Job Details/Information"
		  Text 10, 25, 70, 10, "EMPLOYEE NAME:"
		  Text 150, 25, 60, 10, "HOURLY WAGE:"
		  Text 215, 25, 105, 10, "GROSS MONTHLY EARNINGS:"
		  Text 10, 55, 110, 10, "EMPLOYER/BUSINESS NAME:"
		  Text 10, 95, 25, 10, "Notes:"
		EndDialog


		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = clear_job_btn Then
			JOBS_ARRAY(jobs_employee_name, this_jobs) = ""
			JOBS_ARRAY(jobs_hourly_wage, this_jobs) = ""
			JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs) = ""
			JOBS_ARRAY(jobs_employer_name, this_jobs) = ""
			JOBS_ARRAY(jobs_notes, this_jobs) = ""
		End If
	Loop until ButtonPressed = return_btn
	If JOBS_ARRAY(jobs_employee_name, this_jobs) = "Select One..." Then JOBS_ARRAY(jobs_employee_name, this_jobs) = ""

end function

function format_phone_number(phone_variable, format_type)
'This function formats phone numbers to match the specificed format.
	' format_type_options:
	'  (xxx)xxx-xxxx
	'  xxx-xxx-xxxx
	'  xxx xxx xxxx
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) = 10 Then
		left_phone = left(phone_variable, 3)
		mid_phone = mid(phone_variable, 4, 3)
		right_phone = right(phone_variable, 4)
		format_type = lcase(format_type)
		If format_type = "(xxx)xxx-xxxx" Then
			phone_variable = "(" & left_phone & ")" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx-xxx-xxxx" Then
			phone_variable = left_phone & "-" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx xxx xxxx" Then
			phone_variable = left_phone & " " & mid_phone & " " & right_phone
		End If
	Else
		phone_variable = original_phone_var
	End If
end function

function validate_phone_number(err_msg_variable, list_delimiter, phone_variable, allow_to_be_blank)
'This isn't working yet
'This function will review to ensure a variale appears to be a phone number.
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) <> 10 Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " Phone numbers should be entered as a 10 digit number. Please incldue the area code or check the number to ensure the correct information is entered."
	If len(phone_variable) = 0 then
		If allow_to_be_blank = TRUE then err_msg_variable = ""
	End If
	phone_variable = original_phone_var
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
const memb_notes                    = 53

const jobs_employee_name 			= 0
const jobs_hourly_wage 				= 1
const jobs_gross_monthly_earnings	= 2
const jobs_employer_name 			= 3
const jobs_edit_btn					= 4
const jobs_notes 					= 5

Const end_of_doc = 6			'This is for word document ennumeration

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
Dim JOBS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)
ReDim JOBS_ARRAY(jobs_notes, 0)

'These are all the definitions for droplists
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

'Dimming all the variables because they are defined and set within functions
Dim who_are_we_completing_the_form_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, exp_pay_heat_checkbox, exp_pay_ac_checkbox, exp_pay_electricity_checkbox, exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_pne_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county, caf_form_date, all_the_clients, err_msg

Dim question_1_yn, question_1_notes, question_1_verif_yn, question_1_verif_details
Dim question_2_yn, question_2_notes, question_2_verif_yn, question_2_verif_details
Dim question_3_yn, question_3_notes, question_3_verif_yn, question_3_verif_details
Dim question_4_yn, question_4_notes, question_4_verif_yn, question_4_verif_details
Dim question_5_yn, question_5_notes, question_5_verif_yn, question_5_verif_details
Dim question_6_yn, question_6_notes, question_6_verif_yn, question_6_verif_details
Dim question_7_yn, question_7_notes, question_7_verif_yn, question_7_verif_details
Dim question_8_yn, question_8a_yn, question_8_notes, question_8_verif_yn, question_8_verif_details
Dim question_9_yn, question_9_notes, question_9_verif_yn, question_9_verif_details
Dim question_10_yn, question_10_notes, question_10_verif_yn, question_10_verif_details, question_10_monthly_earnings
Dim question_11_yn, question_11_notes, question_11_verif_yn, question_11_verif_details
Dim pwe_selection
Dim question_12_yn, question_12_notes, question_12_verif_yn, question_12_verif_details
Dim question_12_rsdi_yn, question_12_rsdi_amt, question_12_ssi_yn, question_12_ssi_amt,  question_12_va_yn, question_12_va_amt, question_12_ui_yn, question_12_ui_amt, question_12_wc_yn, question_12_wc_amt, question_12_ret_yn, question_12_ret_amt, question_12_trib_yn, question_12_trib_amt, question_12_cs_yn, question_12_cs_amt, question_12_other_yn, question_12_other_amt
Dim question_13_yn, question_13_notes, question_13_verif_yn, question_13_verif_details
Dim question_14_yn, question_14_notes, question_14_verif_yn, question_14_verif_details
Dim question_14_rent_yn, question_14_subsidy_yn, question_14_mortgage_yn, question_14_association_yn, question_14_insurance_yn, question_14_room_yn, question_14_taxes_yn
Dim question_15_yn, question_15_notes, question_15_verif_yn, question_15_verif_details
Dim question_15_heat_ac_yn, question_15_electricity_yn, question_15_cooking_fuel_yn, question_15_water_and_sewer_yn, question_15_garbage_yn, question_15_phone_yn, question_15_liheap_yn
Dim question_16_yn, question_16_notes, question_16_verif_yn, question_16_verif_details
Dim question_17_yn, question_17_notes, question_17_verif_yn, question_17_verif_details
Dim question_18_yn, question_18_notes, question_18_verif_yn, question_18_verif_details
Dim question_19_yn, question_19_notes, question_19_verif_yn, question_19_verif_details
Dim question_20_yn, question_20_notes, question_20_verif_yn, question_20_verif_details
Dim question_20_cash_yn, question_20_acct_yn, question_20_secu_yn, question_20_cars_yn
Dim question_21_yn, question_21_notes, question_21_verif_yn, question_21_verif_details
Dim question_22_yn, question_22_notes, question_22_verif_yn, question_22_verif_details
Dim question_23_yn, question_23_notes, question_23_verif_yn, question_23_verif_details
Dim question_24_yn, question_24_notes, question_24_verif_yn, question_24_verif_details
Dim question_24_rep_payee_yn, question_24_guardian_fees_yn, question_24_special_diet_yn, question_24_high_housing_yn
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_there, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call check_for_MAXIS(true)
Call MAXIS_case_number_finder(MAXIS_case_number)
application_date = date & ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 235, "Interview Script Case number dialog"
  EditBox 105, 90, 60, 15, MAXIS_case_number
  EditBox 105, 110, 50, 15, CAF_datestamp
  DropListBox 105, 130, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  CheckBox 110, 165, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 150, 165, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 190, 165, 35, 10, "EMER", EMER_on_CAF_checkbox
  ButtonGroup ButtonPressed
    OkButton 260, 215, 50, 15
    CancelButton 315, 215, 50, 15
    PushButton 125, 215, 15, 15, "!", tips_and_tricks_button
  Text 10, 10, 360, 10, "Start this script at the beginning of the interview and keep it running during the entire course of the interview."
  Text 10, 20, 60, 10, "This script will:"
  Text 20, 30, 170, 10, "- Guide you through all of the interview questions."
  Text 20, 40, 170, 10, "- Capture client answers for CASE:NOTE"
  Text 20, 50, 260, 10, "- Create a document of the interview answers to be saved in the ECF Case File."
  Text 20, 60, 245, 10, "- Provide verbiage guidance for consistent resident interview experience."
  Text 20, 70, 260, 10, "- Store the interview date, time, and legth in a database (an FNS requirement)."
  Text 50, 95, 50, 10, "Case number:"
  Text 10, 115, 90, 10, "Date Application Received:"
  Text 40, 135, 60, 10, "Actual CAF Form:"
  GroupBox 105, 150, 125, 30, "Programs marked on CAF"
  Text 145, 220, 105, 10, "Look for me for Tips and Tricks!"
  Text 15, 185, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
  DropListBox 25, 195, 295, 45, "Alert at the time you attempt to leave the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
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
Call restore_your_work(vars_filled)			'looking for a 'restart' run
Call convert_date_into_MAXIS_footer_month(application_date, MAXIS_footer_month, MAXIS_footer_year)
caf_form_date = application_date			'The CAF form date is defaulted to the application date
If vars_filled = TRUE Then show_known_addr = TRUE		'This is a setting for the address dialog to see the view

'If we already know the variables because we used 'restore your work' OR if there is no case number, we don't need to read the information from MAXIS
If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then

	Call back_to_SELF

	Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)			'Then we create an array of the the full hh list for looping purpoases
	full_hh_list = Split(list_for_array, chr(9))

	Call back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "MEMB")								'Going to MEMB for each person in MAXIS to read the known information.

	'Now we start filling in the full client array for use in the dialogs
	member_counter = 0		'this increments to let us add people to the array depending on the case
	Do
		EMReadScreen clt_ref_nbr, 2, 4, 33										'REading each person
		EMReadScreen clt_last_name, 25, 6, 30
		EMReadScreen clt_first_name, 12, 6, 63
		EMReadScreen clt_age, 3, 8, 76

		ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, member_counter)			'creating and adding to the array as we read all the people
		ALL_CLIENTS_ARRAY(memb_ref_numb, member_counter) = clt_ref_nbr
		ALL_CLIENTS_ARRAY(memb_last_name, member_counter) = replace(clt_last_name, "_", "")
		ALL_CLIENTS_ARRAY(memb_first_name, member_counter) = replace(clt_first_name, "_", "")
		ALL_CLIENTS_ARRAY(memb_age, member_counter) = trim(clt_age)

		member_counter = member_counter + 1
		transmit
		EMReadScreen last_memb, 7, 24, 2
	Loop until last_memb = "ENTER A"

	Call back_to_SELF		'backing out to reset
	Call navigate_to_MAXIS_screen("STAT", "MEMB")	'going back to MEMB
	For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)							'Looping through all of the people and gathering detail for each person.
	    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76		'going to the right member
	    transmit

	    EMReadScreen clt_id_verif, 2, 9, 68										'Reading all the information from MEMB for each person
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

		'Formatting all of the detail found and saving to the array
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

	Call navigate_to_MAXIS_screen("STAT", "MEMI")								'Now fgoing to MEMI to get all the detail about each person from MEMI
	For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)							'Looping through all of the people
	    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76		'Going to the MEMI for this person
	    transmit

	    EMReadScreen clt_mar_status, 1, 7, 40									'Reading all the MEMI detail
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

		'Formatting the information read from the panel into the array
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


'Giving the buttons specific unumerations so they don't think they are eachother
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
add_verif_8_btn				= 1070
add_verif_9_btn				= 1071
add_verif_10_btn			= 1072
add_verif_11_btn			= 1073
add_verif_12_btn			= 1074
add_verif_12_btn			= 1075
add_job_btn					= 1076
add_verif_14_btn			= 1080
add_verif_15_btn			= 1081
add_verif_16_btn			= 1082
add_verif_17_btn			= 1083
add_verif_18_btn			= 1084
add_verif_19_btn			= 1085
add_verif_20_btn			= 1090
add_verif_21_btn			= 1091
add_verif_22_btn			= 1092
add_verif_23_btn			= 1093
add_verif_24_btn			= 1094
clear_job_btn				= 1100
open_r_and_r_button 		= 1200

'Presetting booleans for the dialog looping
show_caf_pg_1_pers_dlg = TRUE
show_caf_pg_1_addr_dlg = TRUE
show_caf_pg_2_hhcomp_dlg = TRUE
show_caf_pg_3_hhinfo_dlg = TRUE
show_caf_pg_4_income_dlg = TRUE
show_caf_pg_5_expenses_dlg = TRUE
show_caf_pg_6_other_dlg = TRUE
show_caf_qual_questions_dlg = TRUE
show_caf_sig_dlg = TRUE

caf_pg_1_pers_dlg_cleared = FALSE
caf_pg_1_addr_dlg_cleared = FALSE
caf_pg_2_hhcomp_dlg_cleared = FALSE
caf_pg_3_hhinfo_dlg_cleared = FALSE
caf_pg_4_income_dlg_cleared = FALSE
caf_pg_5_expenses_dlg_cleared = FALSE
caf_pg_6_other_dlg_cleared = FALSE
caf_caf_qual_questions_dlg_cleared = FALSE
caf_sig_dlg_cleared = FALSE

'This is where all of the main dialogs are called.
'They loop together so that you can move between all of the different dialogs.
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
									If caf_pg_1_pers_dlg_cleared = FALSE Then show_caf_pg_1_pers_dlg = TRUE
									If caf_pg_1_addr_dlg_cleared = FALSE Then show_caf_pg_1_addr_dlg = TRUE
									If caf_pg_2_hhcomp_dlg_cleared = FALSE Then show_caf_pg_2_hhcomp_dlg = TRUE
									If caf_pg_3_hhinfo_dlg_cleared = FALSE Then show_caf_pg_3_hhinfo_dlg = TRUE
									If caf_pg_4_income_dlg_cleared = FALSE Then show_caf_pg_4_income_dlg = TRUE
									If caf_pg_5_expenses_dlg_cleared = FALSE Then show_caf_pg_5_expenses_dlg = TRUE
									If caf_pg_6_other_dlg_cleared = FALSE Then show_caf_pg_6_other_dlg = TRUE
									If caf_caf_qual_questions_dlg_cleared = FALSE Then show_caf_qual_questions_dlg = TRUE

									If caf_sig_dlg_cleared = FALSE Then show_caf_sig_dlg = TRUE

									If show_caf_pg_1_pers_dlg = TRUE Then Call dlg_page_one_pers_and_exp

								Loop until show_caf_pg_1_pers_dlg = FALSE
								save_your_work
								If show_caf_pg_1_addr_dlg = TRUE Then Call dlg_page_one_address
							Loop until show_caf_pg_1_addr_dlg = FALSE
							save_your_work
							If show_caf_pg_2_hhcomp_dlg = TRUE Then Call dlg_page_two_household_comp
						Loop until show_caf_pg_2_hhcomp_dlg = FALSE
						save_your_work
						If show_caf_pg_3_hhinfo_dlg = TRUE Then Call dlg_page_three_household_info
					Loop until show_caf_pg_3_hhinfo_dlg = FALSE
					save_your_work
					If show_caf_pg_4_income_dlg = TRUE Then Call dlg_page_four_income
				Loop until show_caf_pg_4_income_dlg = FALSE
				save_your_work
				If show_caf_pg_5_expenses_dlg = TRUE Then Call dlg_page_five_expenses
			Loop until show_caf_pg_5_expenses_dlg = FALSE
			save_your_work
			If show_caf_pg_6_other_dlg = TRUE Then Call dlg_page_six_assets_and_other
		Loop until show_caf_pg_6_other_dlg = FALSE
		save_your_work
		If show_caf_qual_questions_dlg = TRUE Then Call dlg_qualifying_questions
	Loop until show_caf_qual_questions_dlg = FALSE
	save_your_work
	If show_caf_sig_dlg = TRUE Then Call dlg_signature
Loop until show_caf_sig_dlg = FALSE
save_your_work

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
objSelection.TypeText "Date Completed: " & caf_form_date & vbCR
objSelection.TypeText "DATE OF APPLICATION: " & application_date & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR
objSelection.TypeText "Completed over the phone with: " & who_are_we_completing_the_form_with & vbCR

'Program CAF Information
caf_progs = ""
for each_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
	If ALL_CLIENTS_ARRAY(clt_snap_checkbox, each_memb) = checked AND InStr(caf_progs, "SNAP") = 0 Then caf_progs = caf_progs & ", SNAP"
	If ALL_CLIENTS_ARRAY(clt_cash_checkbox, each_memb) = checked AND InStr(caf_progs, "Cash") = 0 Then caf_progs = caf_progs & ", Cash"
	If ALL_CLIENTS_ARRAY(clt_emer_checkbox, each_memb) = checked AND InStr(caf_progs, "EMER") = 0 Then caf_progs = caf_progs & ", EMER"
Next
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
objDoc.Tables.Add objRange, 16, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objPers1Table = objDoc.Tables(1)		'Creates the table with the specific index'
'This table will be formatted to look similar to the structure of CAF Page 1

objPers1Table.AutoFormat(16)							'This adds the borders to the table and formats it
objPers1Table.Columns(1).Width = 500					'This sets the width of the table.

for row = 1 to 15 Step 2
	objPers1Table.Cell(row, 1).SetHeight 10, 2			'setting the heights of the rows
Next
for row = 2 to 16 Step 2
	objPers1Table.Cell(row, 1).SetHeight 15, 2
Next

'Now we are going to look at the the first and second rows. These have 4 cells to add details in and we will split the row into those 4 then resize them
For row = 1 to 2
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 140, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 85, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
'Now going to each cell and setting teh font size
For col = 1 to 4
	objPers1Table.Cell(1, col).Range.Font.Size = 6
	objPers1Table.Cell(2, col).Range.Font.Size = 12
Next

'Adding the headers
objPers1Table.Cell(1, 1).Range.Text = "APPLICANT'S LEGAL NAME - LAST"
objPers1Table.Cell(1, 2).Range.Text = "FIRST NAME"
objPers1Table.Cell(1, 3).Range.Text = "MIDDLE NAME"
objPers1Table.Cell(1, 4).Range.Text = "OTHER NAMES YOU USE"

'Adding the detail from the dialog
objPers1Table.Cell(2, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_last_name, 0)
objPers1Table.Cell(2, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_first_name, 0)
objPers1Table.Cell(2, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_mid_name, 0)
objPers1Table.Cell(2, 4).Range.Text = ALL_CLIENTS_ARRAY(memb_other_names, 0)

' objPers1Table.Cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleNone			'commented out code dealing with borders
' objPers1Table.Cell(1, 3).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 4).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Borders(wdBorderBottom) = wdLineStyleNone

'Now formatting rows 3 and 4 - 3 is the header and 4 is the actual information
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
'Adding the words to rows 3 and 4
objPers1Table.Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
objPers1Table.Cell(3, 2).Range.Text = "DATE OF BIRTH"
objPers1Table.Cell(3, 3).Range.Text = "GENDER"
objPers1Table.Cell(3, 4).Range.Text = "MARITAL STATUS"

objPers1Table.Cell(4, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_soc_sec_numb, 0)
objPers1Table.Cell(4, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_dob, 0)
objPers1Table.Cell(4, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_gender, 0)
objPers1Table.Cell(4, 4).Range.Text = ALL_CLIENTS_ARRAY(memi_marriage_status, 0)

'Now formatting rows 5 and 6 - 5 is the header and 6 is the actual information
For row = 5 to 6
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(5, col).Range.Font.Size = 6
	objPers1Table.Cell(6, col).Range.Font.Size = 12
Next
'Adding the words to rows 5 and 6
objPers1Table.Cell(5, 1).Range.Text = "ADDRESS WHERE YOU LIVE"
' objPers1Table.Cell(5, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(5, 2).Range.Text = "CITY"
objPers1Table.Cell(5, 3).Range.Text = "STATE"
objPers1Table.Cell(5, 4).Range.Text = "ZIP CODE"

objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full
' objPers1Table.Cell(6, 2).Range.Text = ""
objPers1Table.Cell(6, 2).Range.Text = resi_addr_city
objPers1Table.Cell(6, 3).Range.Text = LEFT(resi_addr_state, 2)
objPers1Table.Cell(6, 4).Range.Text = resi_addr_zip

'Now formatting rows 7 and 8 - 7 is the header and 8 is the actual information
For row = 7 to 8
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(7, col).Range.Font.Size = 6
	objPers1Table.Cell(8, col).Range.Font.Size = 12
Next
'Adding the words to rows 7 and 8
objPers1Table.Cell(7, 1).Range.Text = "MAILING ADDRESS"
' objPers1Table.Cell(7, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(7, 2).Range.Text = "CITY"
objPers1Table.Cell(7, 3).Range.Text = "STATE"
objPers1Table.Cell(7, 4).Range.Text = "ZIP CODE"

objPers1Table.Cell(8, 1).Range.Text = mail_addr_street_full
' objPers1Table.Cell(8, 2).Range.Text = ""
objPers1Table.Cell(8, 2).Range.Text = mail_addr_city
objPers1Table.Cell(8, 3).Range.Text = LEFT(mail_addr_state, 2)
objPers1Table.Cell(8, 4).Range.Text = mail_addr_zip

'Now formatting rows 9 and 10 - 9 is the header and 10 is the actual information
For row = 9 to 10
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 105, 2
	objPers1Table.Cell(row, 2).SetWidth 105, 2
	objPers1Table.Cell(row, 3).SetWidth 105, 2
	objPers1Table.Cell(row, 4).SetWidth 185, 2
Next
For col = 1 to 4
	objPers1Table.Cell(9, col).Range.Font.Size = 6
	objPers1Table.Cell(10, col).Range.Font.Size = 11
Next
'Adding the words to rows 9 and 10
objPers1Table.Cell(9, 1).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 2).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 3).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 4).Range.Text = "DO YOU LIVE ON A RESERVATION?"

'formatting the phone numbers so they all match and fit
Call format_phone_number(phone_one_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_two_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_three_number, "xxx-xxx-xxxx")
If phone_pne_type = "" OR phone_pne_type = "Select One..." Then
	phone_one_info = phone_one_number
Else
	phone_one_info = phone_one_number & " (" & left(phone_pne_type, 1) & ")"
End If

If phone_two_type = "" OR phone_two_type = "Select One..." Then
	phone_two_info = phone_two_number
Else
	phone_two_info = phone_two_number & " (" & left(phone_two_type, 1) & ")"
End If
If phone_three_type = "" OR phone_three_type = "Select One..." Then
	phone_three_info = phone_three_number
Else
	phone_three_info = phone_three_number & " (" & left(phone_three_type, 1) & ")"
End If
objPers1Table.Cell(10, 1).Range.Text = phone_one_info
objPers1Table.Cell(10, 2).Range.Text = phone_two_info
objPers1Table.Cell(10, 3).Range.Text = phone_three_info
objPers1Table.Cell(10, 4).Range.Text = reservation_yn & " - " & reservation_name

'Now formatting rows 11 and 12 - 11 is the header and 12 is the actual information
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
'Adding the words to rows 11 and 12
objPers1Table.Cell(11, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
objPers1Table.Cell(11, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
objPers1Table.Cell(11, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

objPers1Table.Cell(12, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_interpreter, 0)
objPers1Table.Cell(12, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_spoken_language, 0)
objPers1Table.Cell(12, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_written_language, 0)

'Now formatting rows 13 and 14 - 13 is the header and 14 is the actual information
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
'Adding the words to rows 13 and 14
objPers1Table.Cell(13, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
objPers1Table.Cell(13, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
objPers1Table.Cell(13, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

objPers1Table.Cell(14, 1).Range.Text = ALL_CLIENTS_ARRAY(memi_last_grade, 0)
objPers1Table.Cell(14, 2).Range.Text = "Date: " & ALL_CLIENTS_ARRAY(memi_MN_entry_date, 0) & "   From: " & ALL_CLIENTS_ARRAY(memi_former_state, 0)
objPers1Table.Cell(14, 3).Range.Text = ALL_CLIENTS_ARRAY(memi_citizen, 0)

'Now formatting rows 15 and 16 - 15 is the header and 16 is the actual information
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
'Adding the words to rows 15 and 16
objPers1Table.Cell(15, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
objPers1Table.Cell(15, 2).Range.Text = "ETHNICITY"
objPers1Table.Cell(15, 3).Range.Text = "RACE"

'defining a string that lists the programs based on the checkboxes of programs from the dialog'
If ALL_CLIENTS_ARRAY(clt_none_checkbox, 0) = checked then progs_applying_for = "NONE"
If ALL_CLIENTS_ARRAY(clt_snap_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", SNAP"
If ALL_CLIENTS_ARRAY(clt_cash_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Cash"
If ALL_CLIENTS_ARRAY(clt_emer_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

'defining a string of the races that were selected from checkboxes in the dialog.
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

objSelection.TypeText "NOTES: " & ALL_CLIENTS_ARRAY(memb_notes, 0) & vbCR

objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF 1 - EXPEDITED QUESTIONS"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
set objEXPTable = objDoc.Tables(2)		'Creates the table with the specific index'

objEXPTable.AutoFormat(16)							'This adds the borders to the table and formats it
objEXPTable.Columns(1).Width = 375					'Setting the widths of the columns
objEXPTable.Columns(2).Width = 120
for col = 1 to 2
	for row = 1 to 8
		objEXPTable.Cell(row, col).Range.Font.Bold = TRUE	'Making the cell text bold.
	next
next

'Adding the Expedited text to the table for Expedited
objEXPTable.Cell(1, 1).Range.Text = "1. How much income (cash or checks) did or will your household get this month?"
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

objEXPTable.Rows(7).Cells.Split 1, 6, TRUE										'Splitting the cells to add more detail for the three questions here
objEXPTable.Cell(7, 1).Range.Text = "When?"
objEXPTable.Cell(7, 2).Range.Text = exp_previous_assistance_when
objEXPTable.Cell(7, 3).Range.Text = "Where?"
objEXPTable.Cell(7, 4).Range.Text = exp_previous_assistance_where
objEXPTable.Cell(7, 5).Range.Text = "What?"
objEXPTable.Cell(7, 6).Range.Text = exp_previous_assistance_what

objEXPTable.Cell(8, 1).Range.Text = "6. Is anyone in your household pregnant?"
If exp_pregnant_who <> "" Then
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn & ", " &  exp_pregnant_who
Else
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn
End If

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.Font.Bold = TRUE
objSelection.TypeText "AGENCY USE:" & vbCr
objSelection.Font.Bold = FALSE
objSelection.TypeText chr(9) & "Intends to reside in MN? - " & ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, 0) & vbCr
objSelection.TypeText chr(9) & "Has Sponsor? - " & ALL_CLIENTS_ARRAY(clt_sponsor_yn, 0) & vbCr
objSelection.TypeText chr(9) & "Immigration Status: " & ALL_CLIENTS_ARRAY(clt_imig_status, 0) & vbCr
objSelection.TypeText chr(9) & "Verification: " & ALL_CLIENTS_ARRAY(clt_verif_yn, 0) & vbCr
If ALL_CLIENTS_ARRAY(clt_verif_details, 0) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & ALL_CLIENTS_ARRAY(clt_verif_details, 0) & vbCr

'Now we have a dynamic number of tables
'each table has to be defined with its index so we need to have a variable to increment
table_count = 3			'table index variable
If UBound(ALL_CLIENTS_ARRAY, 2) <> 0 Then
	ReDim TABLE_ARRAY(UBound(ALL_CLIENTS_ARRAY, 2)-1)		'defining the table array for as many persons aas are in the household - each person gets their own table
	array_counters = 0		'the incrementer for the table array'

	For each_member = 1 to UBound(ALL_CLIENTS_ARRAY, 2)
		objSelection.TypeText "PERSON " & each_member + 1
		Set objRange = objSelection.Range										'range is needed to create tables
		objDoc.Tables.Add objRange, 10, 1										'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)			'Creates the table with the specific index - using the vairable index
		table_count = table_count + 1											'incrementing the table index'

		'This table is now formatted to match how the CAF looks with person information.
		'This formatting uses 'spliting' and resizing to make theym look like the CAF
		TABLE_ARRAY(array_counters).AutoFormat(16)								'This adds the borders to the table and formats it
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

		For row = 3 to 4
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 5, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 95, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 80, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 65, 2
			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
			TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 70, 2
		Next
		For col = 1 to 5
			TABLE_ARRAY(array_counters).Cell(3, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(4, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
		TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "DATE OF BIRTH"
		TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "GENDER"
		TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "RELATIONSHIP TO YOU"
		TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "MARITAL STATUS"

		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_soc_sec_numb, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_dob, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_gender, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = ALL_CLIENTS_ARRAY(memb_rel_to_applct, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = Left(ALL_CLIENTS_ARRAY(memi_marriage_status, each_member), 1)

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
		TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
		TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = ALL_CLIENTS_ARRAY(memb_interpreter, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 2).Range.Text = ALL_CLIENTS_ARRAY(memb_spoken_language, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 3).Range.Text = ALL_CLIENTS_ARRAY(memb_written_language, each_member)

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

		progs_applying_for = ""
		If ALL_CLIENTS_ARRAY(clt_none_checkbox, each_member) = checked then progs_applying_for = "NONE"
		If ALL_CLIENTS_ARRAY(clt_snap_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", SNAP"
		If ALL_CLIENTS_ARRAY(clt_cash_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Cash"
		If ALL_CLIENTS_ARRAY(clt_emer_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
		If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

		race_to_enter = ""
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
	objSelection.TypeText "THERE ARE NO OTHER PEOPLE TO BE LISTED ON THIS APPLICATION" & vbCr
	ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
End If

'This is the rest of the verbiage from the CAF. It is not kept in tables - for the most part
objSelection.TypeText "Q 1. Does everyone in your household buy, fix or eat food with you?" & vbCr
objSelection.TypeText chr(9) & question_1_yn & vbCr
If question_1_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_1_notes & vbCr
If question_1_verif_yn <> "Mot Needed" AND question_1_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_1_verif_yn & vbCr
If question_1_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_1_verif_details & vbCr

objSelection.TypeText "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?" & vbCr
objSelection.TypeText chr(9) & question_2_yn & vbCr
If question_2_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_2_notes & vbCr
If question_2_verif_yn <> "Mot Needed" AND question_2_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_2_verif_yn & vbCr
If question_2_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_2_verif_details & vbCr

objSelection.TypeText "Q 3. Is anyone in the household attending school?" & vbCr
objSelection.TypeText chr(9) & question_3_yn & vbCr
If question_3_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_3_notes & vbCr
If question_3_verif_yn <> "Mot Needed" AND question_3_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_3_verif_yn & vbCr
If question_3_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_3_verif_details & vbCr

objSelection.TypeText "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)" & vbCr
objSelection.TypeText chr(9) & question_4_yn & vbCr
If question_4_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_4_notes & vbCr
If question_4_verif_yn <> "Mot Needed" AND question_4_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_4_verif_yn & vbCr
If question_4_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_4_verif_details & vbCr

objSelection.TypeText "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?" & vbCr
objSelection.TypeText chr(9) & question_5_yn & vbCr
If question_5_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_5_notes & vbCr
If question_5_verif_yn <> "Mot Needed" AND question_5_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_5_verif_yn & vbCr
If question_5_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_5_verif_details & vbCr

objSelection.TypeText "Q 6. Is anyone unable to work for reasons other than illness or disability?" & vbCr
objSelection.TypeText chr(9) & question_6_yn & vbCr
If question_6_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_6_notes & vbCr
If question_6_verif_yn <> "Mot Needed" AND question_6_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_6_verif_yn & vbCr
If question_6_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_6_verif_details & vbCr

objSelection.TypeText "Q 7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?" & vbCr
objSelection.TypeText chr(9) & question_7_yn & vbCr
If question_7_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_7_notes & vbCr
If question_7_verif_yn <> "Mot Needed" AND question_7_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_7_verif_yn & vbCr
If question_7_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_7_verif_details & vbCr

objSelection.TypeText "Q 8. Has anyone in the household had a job or been self-employed in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & question_8_yn & vbCr
objSelection.TypeText "Q 8.a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?" & vbCr
objSelection.TypeText chr(9) & question_8a_yn & vbCr
If question_8_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_8_notes & vbCr
If question_8_verif_yn <> "Mot Needed" AND question_8_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_8_verif_yn & vbCr
If question_8_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_8_verif_details & vbCr

objSelection.TypeText "Q 9. Does anyone in the household have a job or expect to get income from a job this month or next month?" & vbCr

job_added = FALSE
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
		job_added = TRUE

		all_the_tables = UBound(TABLE_ARRAY) + 1
		ReDim Preserve TABLE_ARRAY(all_the_tables)
		Set objRange = objSelection.Range					'range is needed to create tables
		objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		table_count = table_count + 1

		TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 400

		TABLE_ARRAY(array_counters).Cell(1, 1).SetHeight 10, 2
		TABLE_ARRAY(array_counters).Cell(3, 1).SetHeight 10, 2

		TABLE_ARRAY(array_counters).Cell(2, 1).SetHeight 15, 2
		TABLE_ARRAY(array_counters).Cell(4, 1).SetHeight 15, 2

		For row = 1 to 2
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 200, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 90, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Font.Size = 6
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Font.Size = 12

		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "EMPLOYEE NAME"
		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "HOURLY WAGE"
		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "GROSS MONTHLY EARNINGS"
		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = JOBS_ARRAY(jobs_employee_name, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = JOBS_ARRAY(jobs_hourly_wage, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)

		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "EMPLOYER/BUSINESS NAME"
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = JOBS_ARRAY(jobs_employer_name, each_job)

		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		' objSelection.TypeParagraph()						'adds a line between the table and the next information

		array_counters = array_counters + 1

		objSelection.TypeText "NOTES: " & JOBS_ARRAY(jobs_notes, each_job) & vbCR
	End If
next

If job_added = FALSE Then objSelection.TypeText chr(9) & "THERE ARE NO JOBS ENTERED." & vbCr

If question_9_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_9_notes & vbCr
If question_9_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_9_verif_yn & vbCr
If question_9_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_9_verif_details & vbCr

objSelection.TypeText "Q 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?" & vbCr
objSelection.TypeText chr(9) & question_10_yn & vbCr
If question_10_monthly_earnings <> "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: " & question_10_monthly_earnings & vbCr
If question_10_monthly_earnings = "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: NONE LISTED" & vbCr
If question_10_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_10_notes & vbCr
If question_10_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_10_verif_yn & vbCr
If question_10_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_10_verif_details & vbCr

objSelection.TypeText "Q 11. Do you expect any changes in income, expenses or work hours?" & vbCr
objSelection.TypeText chr(9) & question_11_yn & vbCr
If question_11_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_11_notes & vbCr
If question_11_verif_yn <> "Mot Needed" AND question_11_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_11_verif_yn & vbCr
If question_11_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_11_verif_details & vbCr

objSelection.Font.Bold = TRUE
objSelection.TypeText "Principal Wage Earner (PWE)" & vbCr
objSelection.Font.Bold = FALSE

all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 2					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
TABLE_ARRAY(array_counters).Columns(1).Width = 200
TABLE_ARRAY(array_counters).Columns(2).Width = 200

TABLE_ARRAY(array_counters).Cell(1, 1).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(1, 2).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(2, 1).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(2, 2).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(1, 1).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Font.Size = 12
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Font.Size = 12

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text ="DESIGNATED PWE"
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text =pwe_selection
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text ="SIGNATURE OF APPLICANT"
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text ="VERBAL SIGNATURE"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

objSelection.TypeText "Q 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 5, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 500

For row = 1 to 4
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 75, 2
Next
TABLE_ARRAY(array_counters).Rows(5).Cells.Split 1, 3, TRUE

TABLE_ARRAY(array_counters).Cell(5, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(5, 2).SetWidth 175, 2
TABLE_ARRAY(array_counters).Cell(5, 3).SetWidth 75, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_12_rsdi_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "RSDI"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "$ " & question_12_rsdi_amt
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = question_12_ssi_yn
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = "SSI"
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "$ " & question_12_ssi_amt

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_12_va_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Veteran Benefits (VA)"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = "$ " & question_12_va_amt
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = question_12_ui_yn
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = "Unemployment Insurance"
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "$ " & question_12_ui_amt

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_12_wc_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Workers' Compensation"
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "$ " & question_12_wc_amt
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = question_12_ret_yn
TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "Retirement Benefits"
TABLE_ARRAY(array_counters).Cell(3, 6).Range.Text = "$ " & question_12_ret_amt

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_12_trib_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Tribal payments"
TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = "$ " & question_12_trib_amt
TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = question_12_cs_yn
TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = "Child or Spousal support"
TABLE_ARRAY(array_counters).Cell(4, 6).Range.Text = "$ " & question_12_cs_amt

TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = question_12_other_yn
TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "Other unearned income"
TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "$ " & question_12_other_amt

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_12_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_12_notes & vbCr
If question_12_verif_yn <> "Mot Needed" AND question_12_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_12_verif_yn & vbCr
If question_12_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_12_verif_details & vbCr

objSelection.TypeText "Q 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?" & vbCr
objSelection.TypeText chr(9) & question_13_yn & vbCr
If question_13_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_13_notes & vbCr
If question_13_verif_yn <> "Mot Needed" AND question_13_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_13_verif_yn & vbCr
If question_13_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_13_verif_details & vbCr

objSelection.TypeText "Q 14. Does your household have the following housing expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 3
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next
TABLE_ARRAY(array_counters).Rows(4).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(4, 1).SetWidth 90, 2
TABLE_ARRAY(array_counters).Cell(4, 2).SetWidth 430, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_14_rent_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Rent (include mobile home lot rental)"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_14_subsidy_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Rent or Section 8 subsidy"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_14_mortgage_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Mortgage/contract for deed payment"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_14_association_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Association fees"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_14_insurance_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Homeowner's insurance (if not included in mortgage) "
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = question_14_room_yn
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "Room and/or board"

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_14_taxes_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Real estate taxes (if not included in mortgage)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_14_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_14_notes & vbCr
If question_14_verif_yn <> "Mot Needed" AND question_14_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_14_verif_yn & vbCr
If question_14_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_14_verif_details & vbCr

objSelection.TypeText "Q 15. Does your household have the following utility expenses any time during the year?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 3, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 525

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 100, 2
Next
TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(3, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(3, 2).SetWidth 450, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_15_heat_ac_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Heating/air conditioning"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_15_electricity_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Electricity"
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = question_15_cooking_fuel_yn
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "Cooking fuel"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_15_water_and_sewer_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Water and sewer"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_15_garbage_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Garbage removal"
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = question_15_phone_yn
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "Phone/cell phone"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_15_liheap_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_15_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_15_notes & vbCr
If question_15_verif_yn <> "Mot Needed" AND question_15_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_15_verif_yn & vbCr
If question_15_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_15_verif_details & vbCr

objSelection.TypeText "Q 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & question_13_yn & vbCr
If question_16_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_16_notes & vbCr
If question_16_verif_yn <> "Mot Needed" AND question_16_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_16_verif_yn & vbCr
If question_16_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_16_verif_details & vbCr

objSelection.TypeText "Q 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & question_13_yn & vbCr
If question_17_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_17_notes & vbCr
If question_17_verif_yn <> "Mot Needed" AND question_17_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_17_verif_yn & vbCr
If question_17_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_17_verif_details & vbCr

objSelection.TypeText "Q 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?" & vbCr
objSelection.TypeText chr(9) & question_18_yn & vbCr
If question_18_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_18_notes & vbCr
If question_18_verif_yn <> "Mot Needed" AND question_18_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_18_verif_yn & vbCr
If question_18_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_18_verif_details & vbCr

objSelection.TypeText "Q 19. For SNAP only: Does anyone in the household have medical expenses? " & vbCr
objSelection.TypeText chr(9) & question_19_yn & vbCr
If question_19_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_19_notes & vbCr
If question_19_verif_yn <> "Mot Needed" AND question_19_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_19_verif_yn & vbCr
If question_19_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_19_verif_details & vbCr

objSelection.TypeText "Q 20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. " & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_20_cash_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Cash"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_20_acct_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Bank accounts (savings, checking, debit card, etc.)"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_20_secu_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Stocks, bonds, annuities, 401K, etc."
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_20_cars_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Vehicles (cars, trucks, motorcycles, campers, trailers)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_20_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_20_notes & vbCr
If question_20_verif_yn <> "Mot Needed" AND question_20_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_20_verif_yn & vbCr
If question_20_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_20_verif_details & vbCr


objSelection.TypeText "Q 21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)" & vbCr
objSelection.TypeText chr(9) & question_21_yn & vbCr
If question_21_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_21_notes & vbCr
If question_21_verif_yn <> "Mot Needed" AND question_21_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_21_verif_yn & vbCr
If question_21_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_21_verif_details & vbCr

objSelection.TypeText "Q 22. For recertifications only: Did anyone move in or out of your home in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & question_22_yn & vbCr
If question_22_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_22_notes & vbCr
If question_22_verif_yn <> "Mot Needed" AND question_22_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_22_verif_yn & vbCr
If question_22_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_22_verif_details & vbCr

objSelection.TypeText "Q 23. For children under the age of 19, are both parents living in the home?" & vbCr
objSelection.TypeText chr(9) & question_23_yn & vbCr
If question_23_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_23_notes & vbCr
If question_23_verif_yn <> "Mot Needed" AND question_23_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_23_verif_yn & vbCr
If question_23_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_23_verif_details & vbCr

objSelection.TypeText "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_24_rep_payee_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Representative Payee fees"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_24_guardian_fees_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Guardian or Conservator fees"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_24_special_diet_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Physician-prescribed special diet "
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_24_high_housing_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "High housing costs"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_24_notes <> "" Then objSelection.TypeText chr(9) & "Notes: " & question_24_notes & vbCr
If question_24_verif_yn <> "Mot Needed" AND question_24_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_24_verif_yn & vbCr
If question_24_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_24_verif_details & vbCr

objSelection.TypeText "CAF QUALIFYING QUESTIONS" & vbCr

objSelection.TypeText "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?" & vbCr
objSelection.TypeText chr(9) & qual_question_one & vbCr
If trim(qual_memb_one) <> "" AND qual_memb_one <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_one & vbCr
objSelection.TypeText "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?" & vbCr
objSelection.TypeText chr(9) & qual_question_two & vbCr
If trim(qual_memb_two) <> "" AND qual_memb_two <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_two & vbCr
objSelection.TypeText "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?" & vbCr
objSelection.TypeText chr(9) & qual_question_three & vbCr
If trim(qual_memb_there) <> "" AND qual_memb_there <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_there & vbCr
objSelection.TypeText "Has anyone in your household been convicted of a drug felony in the past 10 years?" & vbCr
objSelection.TypeText chr(9) & qual_question_four & vbCr
If trim(qual_memb_four) <> "" AND qual_memb_four <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_four & vbCr
objSelection.TypeText "Is anyone in your household currently violating a condition of parole, probation or supervised release?" & vbCr
objSelection.TypeText chr(9) & qual_question_five & vbCr
If trim(qual_memb_five) <> "" AND qual_memb_five <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_five & vbCr

objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE
objSelection.TypeText "Verbal Signature accepted on " & caf_form_date

' MsgBox "DOC IS CREATED"			'This can be used for testing so we don't add fake documents to the assignment folder.

'Here we are creating the file path and saving the file
file_safe_date = replace(date, "/", "-")		'dates cannot have / for a file name so we change it to a -
'We set the file path and name based on case number and date. We can add other criteria if important.
'This MUST have the 'pdf' file extension to work
If MAXIS_case_number <> "" Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CAF Forms for ECF\CAF - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
If no_case_number_checkbox = checked Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\CAF Forms for ECF\CAF - NEW CASE " & Left(ALL_CLIENTS_ARRAY(memb_first_name, 0), 1) & ". " & ALL_CLIENTS_ARRAY(memb_last_name, 0) & " on " & file_safe_date & ".pdf"
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
	If MAXIS_case_number <> "" Then
		local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	Else
		local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"
	End If

	'we are checking the save your work text file. If it exists we need to delete it because we don't want to save that information locally.
	If objFSO.FileExists(local_changelog_path) = True then
		objFSO.DeleteFile(local_changelog_path)			'DELETE
	End If

	'Now we case note!
	Call start_a_blank_case_note
	Call write_variable_in_CASE_NOTE("CAF Form completed via Phone")
	Call write_variable_in_CASE_NOTE("Form information taken verbally per COVID Waiver Allowance.")
	Call write_variable_in_CASE_NOTE("Form information taken on " & caf_form_date)
	Call write_variable_in_CASE_NOTE("CAF for application date: " & application_date)
	Call write_variable_in_CASE_NOTE("CAF information saved and will be added to ECF within a few days. Detail can be viewed in 'Assignments Folder'.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	'setting the end message
	end_msg = "Success! The information you have provided for the CAF form has been saved to the Assignments forlder so the CAF Form can be updated and added to ECF. The case can be processed using the information saved in the PDF. Additional notes and information are needed or case processing. This script has NOT updated MAXIS or added CAF processing notes."

	'Now we ask if the worker would like the PDF to be opened by the script before the script closes
	'This is helpful because they may not be familiar with where these are saved and they could work from the PDF to process the reVw
	reopen_pdf_doc_msg = MsgBox("The information about the CAF has been saved to a PDF on the LAN to be added to the DHS form and added to ECF." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbYesNo, "Open PDF Doc?")
	If reopen_pdf_doc_msg = vbYes Then
		run_path = chr(34) & pdf_doc_path & chr(34)
		wshshell.Run run_path
		end_msg = end_msg & vbCr & vbCr & "The PDF has been opened for you to view the information that has been saved."
	End If
Else
	end_msg = "Something has gone wrong - the CAF information has NOT been saved correctly to be processed." & vbCr & vbCr & "You can either save the Word Document that has opened as a PDF in the Assignment folder OR Close that document without saving and RERUN the script. Your details have been saved and the script can reopen them and attampt to create the files again. When the script is running, it is best to not interrupt the process."
End If

Call script_end_procedure_with_error_report(end_msg)
