'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - PA VERIF REQUEST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 100                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/20/2021", "Adding GRH functionality to PA Verif Request.##~## ##~##You can now resend WCOM for GRH eligibility and a MEMO for issuance amounts for active or previous GRH eligibility.##~##", "Casey Love, Hennepin County")
call changelog_update("06/23/2021", "PA Verif Request is back!##~## ##~##It is now built with the functionality to either create a SPEC/MEMO of benefit issuances from INQX, or resend a WCOM of an Eligibility Notice.##~## ##~##This new process follows the procedure detailed in the HSR Manual.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "BUG FIX - error for cases with a Significant Change detail in the budget. Added a fix to move past it.", "Casey Love, Hennepin County")
call changelog_update("11/12/2020", "Updated HSR Manual link for Data Privacy due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("07/29/2020", "Removed the 'PRINT' default of the document at the end of the script run because we are not currently in the office.", "Casey Love, Hennepin County")
call changelog_update("07/29/2020", "Removed the option to include income information from MAXIS in the document. The official policy and process needs to be followed for this type of information. Added a button to open the HSR Manual page.", "Casey Love, Hennepin County")
call changelog_update("01/15/2019", "Updated to accomodate benefits larger than $1,000 for SNAP, MFIP, and DWP.", "Casey Love, Hennepin County")
call changelog_update("12/01/2016", "Checkbox added with the option to have 'Other Income' not listed on the word document.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'FUNCTIONS BLOCK ===========================================================================================================
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

        If county_line = "01" Then addr_county = "01 - Aitkin"
        If county_line = "02" Then addr_county = "02 - Anoka"
        If county_line = "03" Then addr_county = "03 - Becker"
        If county_line = "04" Then addr_county = "04 - Beltrami"
        If county_line = "05" Then addr_county = "05 - Benton"
        If county_line = "06" Then addr_county = "06 - Big Stone"
        If county_line = "07" Then addr_county = "07 - Blue Earth"
        If county_line = "08" Then addr_county = "08 - Brown"
        If county_line = "09" Then addr_county = "09 - Carlton"
        If county_line = "10" Then addr_county = "10 - Carver"
        If county_line = "11" Then addr_county = "11 - Cass"
        If county_line = "12" Then addr_county = "12 - Chippewa"
        If county_line = "13" Then addr_county = "13 - Chisago"
        If county_line = "14" Then addr_county = "14 - Clay"
        If county_line = "15" Then addr_county = "15 - Clearwater"
        If county_line = "16" Then addr_county = "16 - Cook"
        If county_line = "17" Then addr_county = "17 - Cottonwood"
        If county_line = "18" Then addr_county = "18 - Crow Wing"
        If county_line = "19" Then addr_county = "19 - Dakota"
        If county_line = "20" Then addr_county = "20 - Dodge"
        If county_line = "21" Then addr_county = "21 - Douglas"
        If county_line = "22" Then addr_county = "22 - Faribault"
        If county_line = "23" Then addr_county = "23 - Fillmore"
        If county_line = "24" Then addr_county = "24 - Freeborn"
        If county_line = "25" Then addr_county = "25 - Goodhue"
        If county_line = "26" Then addr_county = "26 - Grant"
        If county_line = "27" Then addr_county = "27 - Hennepin"
        If county_line = "28" Then addr_county = "28 - Houston"
        If county_line = "29" Then addr_county = "29 - Hubbard"
        If county_line = "30" Then addr_county = "30 - Isanti"
        If county_line = "31" Then addr_county = "31 - Itasca"
        If county_line = "32" Then addr_county = "32 - Jackson"
        If county_line = "33" Then addr_county = "33 - Kanabec"
        If county_line = "34" Then addr_county = "34 - Kandiyohi"
        If county_line = "35" Then addr_county = "35 - Kittson"
        If county_line = "36" Then addr_county = "36 - Koochiching"
        If county_line = "37" Then addr_county = "37 - Lac Qui Parle"
        If county_line = "38" Then addr_county = "38 - Lake"
        If county_line = "39" Then addr_county = "39 - Lake Of Woods"
        If county_line = "40" Then addr_county = "40 - Le Sueur"
        If county_line = "41" Then addr_county = "41 - Lincoln"
        If county_line = "42" Then addr_county = "42 - Lyon"
        If county_line = "43" Then addr_county = "43 - Mcleod"
        If county_line = "44" Then addr_county = "44 - Mahnomen"
        If county_line = "45" Then addr_county = "45 - Marshall"
        If county_line = "46" Then addr_county = "46 - Martin"
        If county_line = "47" Then addr_county = "47 - Meeker"
        If county_line = "48" Then addr_county = "48 - Mille Lacs"
        If county_line = "49" Then addr_county = "49 - Morrison"
        If county_line = "50" Then addr_county = "50 - Mower"
        If county_line = "51" Then addr_county = "51 - Murray"
        If county_line = "52" Then addr_county = "52 - Nicollet"
        If county_line = "53" Then addr_county = "53 - Nobles"
        If county_line = "54" Then addr_county = "54 - Norman"
        If county_line = "55" Then addr_county = "55 - Olmsted"
        If county_line = "56" Then addr_county = "56 - Otter Tail"
        If county_line = "57" Then addr_county = "57 - Pennington"
        If county_line = "58" Then addr_county = "58 - Pine"
        If county_line = "59" Then addr_county = "59 - Pipestone"
        If county_line = "60" Then addr_county = "60 - Polk"
        If county_line = "61" Then addr_county = "61 - Pope"
        If county_line = "62" Then addr_county = "62 - Ramsey"
        If county_line = "63" Then addr_county = "63 - Red Lake"
        If county_line = "64" Then addr_county = "64 - Redwood"
        If county_line = "65" Then addr_county = "65 - Renville"
        If county_line = "66" Then addr_county = "66 - Rice"
        If county_line = "67" Then addr_county = "67 - Rock"
        If county_line = "68" Then addr_county = "68 - Roseau"
        If county_line = "69" Then addr_county = "69 - St. Louis"
        If county_line = "70" Then addr_county = "70 - Scott"
        If county_line = "71" Then addr_county = "71 - Sherburne"
        If county_line = "72" Then addr_county = "72 - Sibley"
        If county_line = "73" Then addr_county = "73 - Stearns"
        If county_line = "74" Then addr_county = "74 - Steele"
        If county_line = "75" Then addr_county = "75 - Stevens"
        If county_line = "76" Then addr_county = "76 - Swift"
        If county_line = "77" Then addr_county = "77 - Todd"
        If county_line = "78" Then addr_county = "78 - Traverse"
        If county_line = "79" Then addr_county = "79 - Wabasha"
        If county_line = "80" Then addr_county = "80 - Wadena"
        If county_line = "81" Then addr_county = "81 - Waseca"
        If county_line = "82" Then addr_county = "82 - Washington"
        If county_line = "83" Then addr_county = "83 - Watonwan"
        If county_line = "84" Then addr_county = "84 - Wilkin"
        If county_line = "85" Then addr_county = "85 - Winona"
        If county_line = "86" Then addr_county = "86 - Wright"
        If county_line = "87" Then addr_county = "87 - Yellow Medicine"
        If county_line = "89" Then addr_county = "89 - Out-of-State"
        resi_county = addr_county

		' Call get_state_name_from_state_code(state_line, resi_state, TRUE)		'This function makes the state code to be the state name written out - including the code
		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If state_line = left(state_item, 2) Then
				resi_state = state_item
			End If
		Next

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

function access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)

	Call navigate_to_MAXIS_screen("STAT", "AREP")

	EMReadScreen arep_name, 37, 4, 32
	arep_name = replace(arep_name, "_", "")
	If arep_name <> "" Then
		EMReadScreen arep_street_one, 22, 5, 32
		EMReadScreen arep_street_two, 22, 6, 32
		EMReadScreen arep_addr_city, 15, 7, 32
		EMReadScreen arep_addr_state, 2, 7, 55
		EMReadScreen arep_addr_zip, 5, 7, 64

		arep_street_one = replace(arep_street_one, "_", "")
		arep_street_two = replace(arep_street_two, "_", "")
		arep_addr_street = arep_street_one & " " & arep_street_two
		arep_addr_street = trim( arep_addr_street)
		arep_addr_city = replace(arep_addr_city, "_", "")
		arep_addr_state = replace(arep_addr_state, "_", "")
		arep_addr_zip = replace(arep_addr_zip, "_", "")

		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If arep_addr_state = left(state_item, 2) Then
				arep_addr_state = state_item
			End If
		Next

		EMReadScreen arep_phone_one, 14, 8, 34
		EMReadScreen arep_ext_one, 3, 8, 55
		EMReadScreen arep_phone_two, 14, 9, 34
		EMReadScreen arep_ext_two, 3, 8, 55

		arep_phone_one = replace(arep_phone_one, ")", "")
		arep_phone_one = replace(arep_phone_one, "  ", "-")
		arep_phone_one = replace(arep_phone_one, " ", "-")
		If arep_phone_one = "___-___-____" Then arep_phone_one = ""

		arep_phone_two = replace(arep_phone_two, ")", "")
		arep_phone_two = replace(arep_phone_two, "  ", "-")
		arep_phone_two = replace(arep_phone_two, " ", "-")
		If arep_phone_two = "___-___-____" Then arep_phone_two = ""

		arep_ext_one = replace(arep_ext_one, "_", "")
		arep_ext_two = replace(arep_ext_two, "_", "")

		EMReadScreen forms_to_arep, 1, 10, 45
		EMReadScreen mmis_mail_to_arep, 1, 10, 77

	End If

end function


function access_SWKR_panel(access_type, swkr_name, swkr_addr_street, swkr_addr_city, swkr_addr_state, swkr_addr_zip, swkr_phone, swkr_ext, notc_to_swkr)

	Call navigate_to_MAXIS_screen("STAT", "SWKR")

	EMReadScreen swkr_name, 35, 6, 32
	swkr_name = replace(swkr_name, "_", "")
	If swkr_name <> "" Then
		EMReadScreen swkr_street_one, 22, 8, 32
		EMReadScreen swqkr_street_two, 22, 9, 32
		EMReadScreen swkr_addr_city, 15, 10, 32
		EMReadScreen swkr_addr_state, 2, 10, 54
		EMReadScreen swkr_addr_zip, 5, 10, 63

		swkr_street_one = replace(swkr_street_one, "_", "")
		swqkr_street_two = replace(swqkr_street_two, "_", "")
		swkr_addr_street = swkr_street_one & " " & swqkr_street_two
		swkr_addr_street = trim(swkr_addr_street)
		swkr_addr_city = replace(swkr_addr_city, "_", "")
		swkr_addr_state = replace(swkr_addr_state, "_", "")
		swkr_addr_zip = replace(swkr_addr_zip, "_", "")

		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If swkr_addr_state = left(state_item, 2) Then
				swkr_addr_state = state_item
			End If
		Next

		EMReadScreen swkr_phone, 14, 12, 34
		EMReadScreen swkr_ext, 4, 12, 54

		swkr_phone = replace(swkr_phone, ")", "")
		swkr_phone = replace(swkr_phone, "  ", "-")
		swkr_phone = replace(swkr_phone, " ", "-")
		If swkr_phone = "___-___-____" Then swkr_phone = ""

		swkr_ext = replace(swkr_ext, "_", "")

		EMReadScreen notc_to_swkr, 1, 15, 63

	End If

end function

function check_if_mmis_in_session(mmis_running, maxis_region)
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		mmis_running = True
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			mmis_running = True
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				mmis_running = False
			ELSE
				mmis_running = True
			END IF
		END IF
	END IF
	If maxis_region = "PRODUCTION" Then EMWriteScreen "1", 2, 15
	If maxis_region = "INQUIRY DB" Then EMWriteScreen "2", 2, 15
	If maxis_region = "TRAINING" Then EMWriteScreen "3", 2, 15
	transmit
end function


Function Create_List_Of_Notices(notice_panel, notices_array, selected_const, information_const, WCOM_row_const, no_notices, specific_prog)
	Erase notices_array
	no_notices = FALSE
	If notice_panel = "WCOM" Then
		wcom_row = 7
		array_counter = 0
		Do
			EMReadScreen notice_prog, 2,  wcom_row, 26
			save_this_notc = True
			If specific_prog <> "" Then
				If notice_prog <> specific_prog Then save_this_notc = False
			End If
			If save_this_notc = True Then
				ReDim Preserve notices_array(3, array_counter)
				EMReadScreen notice_date, 8,  wcom_row, 16
				EMReadScreen notice_info, 31, wcom_row, 30
				EMReadScreen notice_stat, 8,  wcom_row, 71

				notice_date = trim(notice_date)
				notice_prog = trim(notice_prog)
				notice_info = trim(notice_info)
				notice_stat = trim(notice_stat)

				notices_array(selected_const,    array_counter) = unchecked
				notices_array(information_const, array_counter) = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat
				notices_array(WCOM_row_const,    array_counter) = wcom_row

				array_counter = array_counter + 1
			End If
			wcom_row = wcom_row + 1

			EMReadScreen next_notice, 4, wcom_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If

	If notice_panel = "MEMO" Then
		memo_row = 7
		array_counter = 0
		Do
			ReDim Preserve notices_array(3, array_counter)
			EMReadScreen notice_date, 8,  memo_row, 19
			EMReadScreen notice_info, 31, memo_row, 29
			EMReadScreen notice_stat, 8,  memo_row, 67

			notice_date = trim(notice_date)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			notices_array(selected_const,    array_counter) = unchecked
			notices_array(information_const, array_counter) = notice_info & " - " & notice_date & " - Status: " & notice_stat
			notices_array(WCOM_row_const,    array_counter) = memo_row

			array_counter = array_counter + 1
			memo_row = memo_row + 1

			EMReadScreen next_notice, 4, memo_row, 30
			next_notice = trim(next_notice)

		Loop until next_notice = ""
	End If
	If array_counter = 0 Then no_notices = TRUE
End Function

function leave_notice_text(ask_first)
	EMReadScreen notice_indicator, 6, 2, 72

	If notice_indicator = "FMINFO" Then
		If ask_first = True Then ask_to_leave_msg = MsgBox("It appears we are in a notice text, in order to contine, we must leave the notice text." & vbCr & vbCr & "Is it alright to leave to notice text now?", vbQuestion + vbYesNo, "Leave Notice Text")
		If ask_to_leave_msg = vbYes OR ask_first = False Then  PF3
	End If
end function

function Select_New_WCOM(notices_array, selected_const, information_const, WCOM_row_const, case_number_known, allow_wcom, allow_memo, notc_month, notc_year, no_notices, specific_prog, allow_multiple_notc, allow_cancel)
	If allow_wcom = True AND allow_memo = True Then
		notice_panel = "Select One..."
	ElseIf allow_wcom = True Then
		notice_panel = "WCOM"
	ElseIf allow_memo = True Then
		notice_panel = "MEMO"
	End If
	Do
	    Do
	    	err_msg = ""

	    	dlg_y_pos = 85
	    	dlg_length = 145
			If no_notices = False Then dlg_length = dlg_length + (UBound(notices_array, 2) * 20)

	        Dialog1 = ""
	    	BeginDialog Dialog1, 0, 0, 205, dlg_length, "Notices to Print"
	    	  Text 5, 10, 50, 10, "Case Number"
	    	  If case_number_known = False Then EditBox 65, 5, 50, 15, MAXIS_case_number
			  If case_number_known = True Then Text 65, 10, 50, 15, MAXIS_case_number
	    	  Text 5, 30, 130, 10, "Where is the notice you want to print?"
	    	  If allow_wcom = True AND allow_memo = True Then
			      DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"WCOM"+chr(9)+"MEMO", notice_panel
			  ElseIf allow_wcom = True Then
			  	  DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"WCOM", notice_panel
			  ElseIf allow_memo = True Then
			  	  DropListBox 140, 25, 60, 45, "Select One..."+chr(9)+"MEMO", notice_panel
			  End If
	    	  Text 5, 50, 120, 10, "In which month was the notice sent?"
	    	  EditBox 140, 45, 20, 15, notc_month
	    	  EditBox 165, 45, 20, 15, notc_year
	    	  ButtonGroup ButtonPressed
	    	    PushButton 60, 70, 50, 10, "Find Notices", find_notices_button
	    	  If no_notices = FALSE Then
	    		  For notices_listed = 0 to UBound(notices_array, 2)
	    		  	CheckBox 10, dlg_y_pos, 185, 10, notices_array(information_const, notices_listed), notices_array(selected_const, notices_listed)
	    			dlg_y_pos = dlg_y_pos + 15
	    		  Next
	    	  Else
	    	  	  Text 10, dlg_y_pos, 185, 10, "**No Notices could be found here.**"
	    		  dlg_y_pos = dlg_y_pos + 15
	    	  End If
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  If case_number_known = False Then EditBox 75, dlg_y_pos, 125, 15, worker_signature
			  If case_number_known = True Then Text 80, dlg_y_pos, 125, 15, worker_signature
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  Text 5, dlg_y_pos, 60, 10, "Worker Signature:"
	    	  dlg_y_pos = dlg_y_pos + 15
	    	  If allow_cancel = True Then
				  ButtonGroup ButtonPressed
		    	    OkButton 100, dlg_y_pos, 50, 15
		    	    CancelButton 150, dlg_y_pos, 50, 15
			  Else
				  ButtonGroup ButtonPressed
					OkButton 150, dlg_y_pos, 50, 15
			  End If
	    	  dlg_y_pos = dlg_y_pos + 5
	    	  If case_number_known = False Then CheckBox 5, dlg_y_pos, 90, 10, "Check here to case note.", case_note_check
	    	EndDialog

	    	Dialog Dialog1
	    	If allow_cancel = True Then cancel_confirmation

	    	notice_selected = FALSE
			If no_notices = False Then
		    	For notice_to_print = 0 to UBound(notices_array, 2)
		    		If notices_array(selected_const, notice_to_print) = checked Then
						If allow_multiple_notc = False AND notice_selected = TRUE AND InStr(err_msg, "One one NOTICE can be selected.") = 0 Then err_msg = err_msg & vbNewLine &  "- One one NOTICE can be selected."
						notice_selected = TRUE
					End If
		    	Next
			End If

			If case_number_known = False Then
	    		If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "- Enter a Case Number."
			End If
	    	If notice_panel = "Select One..." Then err_msg = err_msg & vbNewLine & "- Select where the notice to print is."
	    	If notc_month = "" or notc_year = "" Then err_msg = err_msg & vbNewLine & "- Enter footer month and year."
	    	If notice_selected = False Then err_msg = err_msg & vbNewLine & "- Select a notice to be copied to a Word Document."


	    	If ButtonPressed = find_notices_button then
	    		If notice_panel <> "Select One..." AND MAXIS_case_number <> "" AND notc_month <> "" AND notc_year <> "" Then
	    			Call navigate_to_MAXIS_screen ("SPEC", notice_panel)
	    			If notice_panel = "MEMO" then
	    				EMWriteScreen notc_month, 3, 48
	    				EMWriteScreen MAXIS_footer_year, 3, 53
	    			ElseIf notice_panel = "WCOM" Then
	    				EMWriteScreen notc_month, 3, 46
	    				EMWriteScreen notc_year, 3, 51
	    			End If
	    			transmit
					Call Create_List_Of_Notices(notice_panel, notices_array, selected_const, information_const, WCOM_row_const, no_notices, specific_prog)
	    			err_msg = "LOOP"
	    		Else
	    			err_msg = err_msg & vbNewLine & "!!! Cannot read a list of notices without a panel selected, a case number entered, and footer month & year entered !!!"
	    		End If
	    	End If

	    	If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg

	    Loop Until err_msg = ""
	    Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
end function

function resend_existing_wcom(wcom_month, wcom_year, wcom_row, wcom_success, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip)

	If search_for_arep_and_swkr = True Then
		call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
		EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
		call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
		EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	End If

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")

	EMWriteScreen wcom_month, 3, 46
	EMWriteScreen wcom_year, 3, 51
	transmit
	EMWriteScreen "A", wcom_row, 13
	transmit

	row = 1
	col = 1
	EMSearch "SOCWKR", row, col
	if row <> 0 Then swkr_row = row

	row = 1
	col = 1
	EMSearch "ALTREP", row, col
	if row <> 0 Then arep_row = row

	row = 1
	col = 1
	EMSearch "OTHER", row, col
	if row <> 0 Then other_row = row

	' MsgBox "AREP - " & arep_row & vbCr & "SWKR - " & swkr_row & vbCr & "OTHER - " & other_row

	EMWriteScreen "X", 5, 12 		'This is the CLIENT Row
	If send_to_other = "Y" Then EMWriteScreen "X", other_row, 12		'This is the OTHER Row - need handling
	If forms_to_arep = "Y" Then EMWriteScreen "X", arep_row, 12 		'AREP
	If forms_to_swkr = "Y" Then EMWriteScreen "X", swkr_row, 12		'SWKR
	transmit
	If send_to_other = "Y" Then
		other_street = trim(other_street)

		EMWriteScreen other_name, 13, 24
		If len(other_street) < 25 Then
			EMWriteScreen other_street, 14, 24
		Else
			other_street_array = split(other_street, " ")
			col = 24
			row = 14
			for each word in other_street_array
				If col + len(word) + 1 > 47 Then
					row = row + 1
					If row = 16 then Exit for
				End If
				EMWriteScreen " " & word, row, col
				col = col + len(word) + 1
			next
		End If
		EMWriteScreen other_city, 16, 24
		EMWriteScreen other_state, 17, 24
		EMWriteScreen other_zip, 17, 32

		transmit
		EMReadScreen post_office_warning, 7, 3, 6
		If post_office_warning = "Warning" Then transmit
	End If
	EMReadScreen recipient_selection_check, 26, 2, 28
	If memo_input_screen = "Notice Recipient Selection" Then transmit
	' MsgBox "Pause and look at the WCOM"

	EMReadScreen check_for_resent, 6, wcom_row, 3
	EMReadScreen check_for_waiting, 7, wcom_row, 71
	' MsgBox "check for resent - " & check_for_resent & vbCr & "check for waiting - " & check_for_waiting
	If check_for_resent = "ReSent" and check_for_waiting = "Waiting" Then wcom_success = True
end function

'END FUNCTIONS BLOCK========================================================================================================

'DECLARATIONS BLOCK========================================================================================================
'Numbering the buttons
snap_change_wcom_btn = 1010
ga_change_wcom_btn   = 1020
msa_change_wcom_btn  = 1030
mfip_change_wcom_btn = 1040
dwp_change_wcom_btn  = 1050
grh_change_wcom_btn  = 1060
hc_change_wcom_btn   = 1070

snap_wcom_btn = 110
ga_wcom_btn   = 120
msa_wcom_btn  = 130
mfip_wcom_btn = 140
dwp_wcom_btn  = 150
grh_wcom_btn  = 160
hc_wcom_btn   = 170

snap_view_inqx_btn = 510
ga_view_inqx_btn   = 520
msa_view_inqx_btn  = 530
mfip_view_inqx_btn = 540
dwp_view_inqx_btn  = 550
grh_view_inqx_btn  = 560
' hc_wcom_btn   = 170

snap_program_history_button = 51
ga_program_history_button 	= 52
msa_program_history_button 	= 53
mfip_program_history_button = 54
dwp_program_history_button 	= 55
grh_program_history_button 	= 56
hc_program_history_button 	= 57

CURR_button = 5001
PERS_button = 5002
NOTE_button = 5003
XFER_button = 5004
WCOM_button = 5005
MEMO_button = 5006
PROG_button = 5007
MEMB_button = 5008
REVW_button = 5009
INQB_button = 5010
INQD_button = 5011
INQX_button = 5012
ELIG_FS_button = 5013
ELIG_MFIP_button = 5014
ELIG_DWP_button = 5015
ELIG_GA_button = 5016
ELIG_MSA_button = 5017
ELIG_GRH_button = 5018
ELIG_HC_button = 5019
ELIG_SUMM_button = 5020
ELIG_DENY_button = 5021

Dim notices_array()
ReDim notices_array(3,0)

Const selected = 0
Const information = 1
Const WCOM_search_row = 2

const cash_grant_amount_const 	= 0
const snap_grant_amount_const 	= 1
const benefit_month_const		= 2
const note_message_const		= 3
const benefit_month_as_date_const = 4
const last_const				= 5

Dim SNAP_ISSUANCE_ARRAY()
ReDim SNAP_ISSUANCE_ARRAY(last_const, 0)

Dim GA_ISSUANCE_ARRAY()
ReDim GA_ISSUANCE_ARRAY(last_const, 0)

Dim MSA_ISSUANCE_ARRAY()
ReDim MSA_ISSUANCE_ARRAY(last_const, 0)

Dim MFIP_ISSUANCE_ARRAY()
ReDim MFIP_ISSUANCE_ARRAY(last_const, 0)

Dim DWP_ISSUANCE_ARRAY()
ReDim DWP_ISSUANCE_ARRAY(last_const, 0)

Dim GRH_ISSUANCE_ARRAY()
ReDim GRH_ISSUANCE_ARRAY(last_const, 0)
'END DECLARATIONS BLOCK========================================================================================================

'THE SCRIPT=================================================================================================================
script_run_lowdown = ""		'setting the script run lowdown for output to a error report email.
EMConnect ""				'connecting to MAXIS

Call MAXIS_case_number_finder(MAXIS_case_number)		'getting the case number
MAXIS_footer_month = CM_plus_1_mo						'setting the footer month to something - this isn't particularly necessary - but we do want it not too old.
MAXIS_footer_year = CM_plus_1_yr
clt_in_person = FALSE									'defaulting if clt is in person
check_for_MAXIS(False)									'make sure we are in MAXIS

'Here we get started
'Initial dialog to start the run - gathering the CASE Number, what kind of request it is, and the worker signature
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 301, 130, "Verification of Public Assistance"
		  EditBox 85, 50, 60, 15, MAXIS_case_number
		  'COMMENTED OUT because we don't have 'in person' functionality readyir heavily needed
		  ' DropListBox 85, 70, 210, 45, "Resident on the Phone (or AREP)"+chr(9)+"Resident in Person (or AREP)"+chr(9)+"Resend TAX Notice of Cash Benefit"+chr(9)+"PHA (Public Housing form)"+chr(9)+"Request of Medical Payment History (from Resident or AREP)"+chr(9)+"Documents from ECF", contact_type
		  DropListBox 85, 70, 210, 45, "Resident on the Phone (or AREP)"+chr(9)+"PHA (Public Housing form)"+chr(9)+"Request of Medical Payment History (from Resident or AREP)"+chr(9)+"Documents from ECF", contact_type
		  EditBox 85, 90, 210, 15, worker_signature
		  ButtonGroup ButtonPressed
		    OkButton 195, 110, 50, 15
		    CancelButton 245, 110, 50, 15
		    PushButton 120, 30, 175, 15, "HSR Manual for Verification of Public Assistance", verif_pa_hsr_manual_btn
		  Text 10, 10, 290, 20, "PA Verif Request script will assist you in following the procedure for Client Requests of their Public Assistance benefits. Details of this process can be found in the HSR Manual."
		  Text 30, 55, 50, 10, "Case Number:"
		  Text 15, 75, 65, 10, "Source of Request:"
		  Text 20, 95, 60, 10, "Worker Signature:"
		EndDialog

		dialog Dialog1
		cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")
		If ButtonPressed = verif_pa_hsr_manual_btn Then
			'opening the HSR manual page for Verification of Public Assitance
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Verification-of-public-assistance.aspx"
			err_msg = "LOOP"
		End If

		If err_msg <> "LOOP" and err_msg <> "" Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg		'showing any errors
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

script_run_lowdown = script_run_lowdown & vbCr & "Contact Type Selected - " & contact_type			'saving information for error output email

'Opening procedural references and ending the script for options that work outside of MAXIS.
If contact_type = "PHA (Public Housing form)" Then
	run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Verification-of-public-assistance.aspx"
	end_msg = "Requests from PHA (Public Housing Agency) of a residents Cash Assistance have a special process." & vbCr & vbCr &_
			  "THESE ARE HANDLED BY Hazel Haynes and Tammy Richert." & vbCr & "---------------------------------------" & vbCr &_
			  "The script has opened the correct HSR Manual page in Edge and you can view the specific procedure under the header 'Cash (requested by Public Housing Agency)'." & vbCr & vbCr &_
			  "You should not use WCOM or MEMO to provide Cash Benefit Verification to PHA." & vbCr &_
			  "For additional questions or clarification, contact Knowledge Now."
	 script_end_procedure(end_msg)
End If
If contact_type = "Request of Medical Payment History (from Resident or AREP)" Then
	run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Verification-of-public-assistance.aspx"
	run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.mn.gov/forms/DHS-2133-ENG"
	end_msg = "Requests of Medical Payment History have a specific DHS Webform (2133) for completion." & vbCr & vbCr &_
			  "THESE ARE NOT HANDLED BY HENNEPIN COUNTY DIRECTLY" & vbCr & "---------------------------------------" & vbCr &_
			  "The script has opened the correct HSR Manual page in Edge as well as the page for the DHS Webform 2133. " & vbCr & vbCr &_
			  "Payment history is kept for 36 months, medical requests can only be made for medical claims paid within three years."
	 script_end_procedure(end_msg)
End If
If contact_type = "Documents from ECF" Then
	run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Data_Privacy.aspx"
	end_msg = "Most requests for documents must be handled by the ROI Team" & vbCr & vbCr &_
			  "The client should contact this team by phone or the data request portal:" & vbCr &_
			  "  - Phone: 612-543-4887" & vbCr &_
			  "  - Online Data Request Portal: www.hennepin.us/datarequest" &vbCr &_
			  "    (this requires users to create a login)" & vbCr & "---------------------------------------" & vbCr &_
			  "Full procedural information from the HSR Manual has opened in Edge for you to review." & vbCr &_
			  "---------------------------------------" & vbCr &_
			  "Once the caller's identity has been verified through other means, you can as an HSR:" & vbCr &_
			  "  - Provide verbal incformation from an ECF Case file for:" &vbCr &_
			  "     - Social Security Card Number" & vbCr &_
			  "     - State Issued ID Number" & vbCr &_
			  "     - Case Number" & vbCr &_
			  "  - Send documents from ECF of:" &vbCr &_
			  "     - Copy of Birth Certificate" & vbCr &_
			  "     - Copy of Government Issued Immigration Document" & vbCr &_
			  "     - Work Number/Equifax Verifications" & vbCr &_
			  "     THESE ARE THE ONLY DOCUMENTS HSRS MAY SEND FROM ECF - WITH NO EXCEPTIONS" & vbCr &_
			  "  - Send Benefit information from MAXIS. (This Script can assist with this process but you must select 'Resident on the Phone' or 'Resident in Person' in the first Dialog.)" &vbCr &_
			  "For additional questions or clarification, contact the ROI Team at HSPH.ROI.POD@hennepin.us."
	 script_end_procedure(end_msg)
End If

Call back_to_SELF						'getting reset in the script run
EMReadScreen MX_region, 12, 22, 48		'ensuring we are not in INQUIRY
MX_region = trim(MX_region)
script_run_lowdown = script_run_lowdown & vbCr & "MAXIS Region - " & MX_region & vbCr			'saving information for error output email
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If
'COMMENTED OUT until we are doing HC things
' Call check_if_mmis_in_session(mmis_running, MX_region)

If contact_type = "Resident in Person (or AREP)" Then clt_in_person = True		'sets this to default to a WORD Document creation once added to the functionality

Call generate_client_list(select_a_client, "Select or Type")					'Making a list for the ComboBox of the HH Members

'Reading what is happening in the case
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
'saving information for error output email
script_run_lowdown = script_run_lowdown & vbCr & "Case information:"
script_run_lowdown = script_run_lowdown & vbCr & "case_active - " & case_active  & vbCr & "case_pending - " & case_pending & vbCr & "case_rein - " & case_rein
script_run_lowdown = script_run_lowdown & vbCr & "family_cash_case - " & family_cash_case & vbCr & "mfip_case - " & mfip_case & vbCr & "dwp_case - " & dwp_case & vbCr & "adult_cash_case - " & adult_cash_case & vbCr & "ga_case - " & ga_case & vbCr & "msa_case - " & msa_case & vbCr & "grh_case - " & grh_case & vbCr & "snap_case - " & snap_case  & vbCr & "ma_case - " & ma_case & vbCr & "msp_case - " & msp_case
script_run_lowdown = script_run_lowdown & vbCr & "unknown_cash_pending - " & unknown_cash_pending & vbCr & "unknown_hc_pending - " & unknown_hc_pending & vbCr & "ga_status - " & ga_status & vbCr & "msa_status - " & msa_status & vbCr & "mfip_status - " & mfip_status & vbCr & "dwp_status - " & dwp_status  & vbCr & "grh_status - " & grh_status & vbCr & "snap_status - " & snap_status  & vbCr & "ma_status - " & ma_status & vbCr & "msp_status - " & msp_status

'Now the script will go to find the benefit amount and the ELIGIBILITY NOTICE from the WCOM of the most recent approved month.
'This is functionality that could possibly be improved on as sometimes we don't find a WCOM

If ga_status = "ACTIVE" Then				'searching for GA Information'

	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23		'find the right program
		If inqb_program = "GA" Then
			EMReadScreen ga_amount, 10, inqb_row, 38	'read the benefit amount listed
			ga_amount = trim(ga_amount)
			Exit Do										'once the first one is found - we're done
		End If
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "  "						'read until the list is done

		Call back_to_SELF		'reset

 	Call navigate_to_MAXIS_screen("ELIG", "GA")			'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 20, 78							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen ga_approved_date, 8, status_row, 26					'reading the date of approval
	ga_approved_date = DateAdd("m", 1, ga_approved_date)				'going to the next month from that date (the plus one from the time of approval
	ga_month = right("00" & DatePart("m", ga_approved_date), 2)			'making the date a footer month and year for this program
	ga_year = right(DatePart("yyyy", ga_approved_date), 2)

	Call back_to_SELF			'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen ga_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen ga_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "GA" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be GA and the title should have ELIG in it.
			ga_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to GA
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			ga_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR ga_wcom_text <> ""
	If ga_wcom_text = "" Then ga_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

If msa_status = "ACTIVE" Then				'searching for MSA Information'
	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23		'find the right program
		If inqb_program = "MS" Then
			EMReadScreen msa_amount, 10, inqb_row, 38	'read the benefit amount listed
			msa_amount = trim(msa_amount)
			Exit Do										'once the first one is found - we're done
		End If
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "  "						'read until the list is done

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("ELIG", "MSA")		'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 20, 79							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen msa_approved_date, 8, status_row, 26					'reading the date of approval
	msa_approved_date = DateAdd("m", 1, msa_approved_date)				'going to the next month from that date (the plus one from the time of approval
	msa_month = right("00" & DatePart("m", msa_approved_date), 2)		'making the date a footer month and year for this program
	msa_year = right(DatePart("yyyy", msa_approved_date), 2)
	transmit

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen msa_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen msa_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "MS" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be MSA and the title should have ELIG in it.
			ga_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to MSA
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			msa_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR msa_wcom_text <> ""
	If msa_wcom_text = "" Then msa_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

If mfip_status = "ACTIVE" Then				'searching for MFIP Information'
	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 5, inqb_row, 23		'find the right program
		If inqb_program = "MF-MF" and mf_mf_amount = "" Then
			EMReadScreen mf_mf_amount, 10, inqb_row, 38	'read the benefit amount listed
			mf_mf_amount = trim(mf_mf_amount)
		End If
		If inqb_program = "MF-FS" and mf_fs_amount = "" Then
			EMReadScreen mf_fs_amount, 10, inqb_row, 38	'read the benefit amount listed
			mf_fs_amount = trim(mf_fs_amount)
		End If
		If inqb_program = "MF-HG" and mf_hg_amount = "" Then
			EMReadScreen mf_hg_amount, 10, inqb_row, 38	'read the benefit amount listed
			mf_hg_amount = trim(mf_hg_amount)
		End If
		If mf_mf_amount <> "" AND mf_fs_amount <> "" AND mf_hg_amount <> "" Then Exit Do
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "     "						'read until the list is done

	If mf_mf_amount <> "" Then mfip_amount = mfip_amount & "CASH: $ " & mf_mf_amount & ", "
	If mf_hg_amount <> "" Then mfip_amount = mfip_amount & "Housing Grant: $ " & mf_hg_amount & ", "
	If mf_fs_amount <> "" Then mfip_amount = mfip_amount & "Food: $ " & mf_fs_amount & ", "
	If right(mfip_amount, 2) = ", " Then mfip_amount = left(mfip_amount, len(mfip_amount) - 2)

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("ELIG", "MFIP")		'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 20, 79							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen mfip_approved_date, 8, status_row, 26					'reading the date of approval
	mfip_approved_date = DateAdd("m", 1, mfip_approved_date)			'going to the next month from that date (the plus one from the time of approval
	mfip_month = right("00" & DatePart("m", mfip_approved_date), 2)		'making the date a footer month and year for this program
	mfip_year = right(DatePart("yyyy", mfip_approved_date), 2)
	transmit

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen mfip_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen mfip_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "MF" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be MFIP and the title should have ELIG in it.
			ga_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to MFIP
			ga_wcom_position = wcom_row - 6

			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			mfip_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR mfip_wcom_text <> ""
	If mfip_wcom_text = "" Then mfip_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

If dwp_status = "ACTIVE" Then
	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23		'find the right program
		If inqb_program = "DW" Then
			EMReadScreen dwp_amount, 10, inqb_row, 38	'read the benefit amount listed
			dwp_amount = trim(dwp_amount)
			Exit Do										'once the first one is found - we're done
		End If
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "  "						'read until the list is done

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("ELIG", "DWP")			'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 19, 78							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen dwp_approved_date, 8, status_row, 26					'reading the date of approval
	dwp_approved_date = DateAdd("m", 1, dwp_approved_date)			'going to the next month from that date (the plus one from the time of approval
	dwp_month = right("00" & DatePart("m", dwp_approved_date), 2)		'making the date a footer month and year for this program
	dwp_year = right(DatePart("yyyy", dwp_approved_date), 2)

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen dwp_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen dwp_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "FS" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be DWP and the title should have ELIG in it.
			dwp_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to DWP
			dwp_wcom_position = wcom_row - 6
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			dwp_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR dwp_wcom_text <> ""
	If dwp_wcom_text = "" Then dwp_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

If snap_status = "ACTIVE" Then				'searching for SNAP Information'
	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23		'find the right program
		If inqb_program = "FS" Then
			EMReadScreen snap_amount, 10, inqb_row, 38	'read the benefit amount listed
			snap_amount = trim(snap_amount)
			Exit Do										'once the first one is found - we're done
		End If
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "  "						'read until the list is done

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("ELIG", "FS")			'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 19, 78							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen snap_approved_date, 8, status_row, 26					'reading the date of approval
	snap_approved_date = DateAdd("m", 1, snap_approved_date)			'going to the next month from that date (the plus one from the time of approval
	snap_month = right("00" & DatePart("m", snap_approved_date), 2)		'making the date a footer month and year for this program
	snap_year = right(DatePart("yyyy", snap_approved_date), 2)

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen snap_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen snap_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "FS" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be SNAP and the title should have ELIG in it.
			snap_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to SNAP
			snap_wcom_position = wcom_row - 6
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			snap_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR snap_wcom_text <> ""
	If snap_wcom_text = "" Then snap_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

If grh_status = "ACTIVE" Then				'searching for GRH Information'
	Call navigate_to_MAXIS_screen("MONY", "INQB")		'reading the recent benefit amount
	inqb_row = 6										'start at the top of the list
	Do
		EMReadScreen inqb_program, 2, inqb_row, 23		'find the right program
		If inqb_program = "GR" Then
			EMReadScreen grh_amount, 10, inqb_row, 38	'read the benefit amount listed
			grh_amount = trim(grh_amount)
			Exit Do										'once the first one is found - we're done
		End If
		inqb_row = inqb_row + 1							'go to the next row
	Loop until inqb_program = "  "						'read until the list is done

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("ELIG", "GRH")			'since we are set to CM + 1, this reads the most recent month
	EMWriteScreen "99", 20, 79							'opening the version histor of ELIG
	transmit

	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		' If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""		'finding the first approved version
	EMReadScreen grh_approved_date, 8, status_row, 26					'reading the date of approval
	grh_approved_date = DateAdd("m", 1, grh_approved_date)			'going to the next month from that date (the plus one from the time of approval
	grh_month = right("00" & DatePart("m", grh_approved_date), 2)		'making the date a footer month and year for this program
	grh_year = right(DatePart("yyyy", grh_approved_date), 2)

	Call back_to_SELF		'reset

	Call navigate_to_MAXIS_screen("SPEC", "WCOM")		'now going to look for a notice
	EMWriteScreen grh_month, 3, 46						'entering the month found from ELIG
	EMWriteScreen grh_year, 3, 51
	transmit

	wcom_row = 7										'looking for a WCOM
	Do
		EMReadScreen prg_typ, 2, wcom_row, 26			'reading the program and title of the notice
		EMReadScreen notc_title, 30, wcom_row, 30

		If prg_typ = "GR" AND InStr(notc_title, "ELIG") <> 0 Then		'the program needs to be SNAP and the title should have ELIG in it.
			grh_wcom_row = wcom_row										'saving the row of the WCOM and which notice in the list it is specific to SNAP
			grh_wcom_position = wcom_row - 6
			EMReadScreen notice_date, 8,  wcom_row, 16
			EMReadScreen notice_prog, 2,  wcom_row, 26
			EMReadScreen notice_info, 31, wcom_row, 30
			EMReadScreen notice_stat, 8,  wcom_row, 71

			notice_date = trim(notice_date)
			notice_prog = trim(notice_prog)
			notice_info = trim(notice_info)
			notice_stat = trim(notice_stat)

			grh_wcom_text = notice_info & " - " & notice_prog & " - " & notice_date & " - Status: " & notice_stat	'this is what is output on the dialog'
		End If
		wcom_row = wcom_row + 1
	Loop until prg_typ = "  " OR grh_wcom_text <> ""
	If grh_wcom_text = "" Then grh_wcom_text = "NO WCOM Found"		'if no WCOM found for this program, in this month, with ELIG in the title, cannot default a WCOM - creating output for the dialog.
End If

'COMMENTED OUT until we write the HC portion
' If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then
' End If

'defaulting the program history information.
snap_prog_history_exists = False
ga_prog_history_exists = False
msa_prog_history_exists = False
mfip_prog_history_exists = False
dwp_prog_history_exists = False
grh_prog_history_exists = False

'Now we are going to search for program history for any program that is not active
'having program history set to true will add that program to the dialog to send a EMMO of benefits.
Call navigate_to_MAXIS_screen("CASE", "CURR")
EMWriteScreen "X", 4, 9			'opens the Program History page
transmit

If snap_status <> "ACTIVE" Then				'Looking for SNAP program history
	EMWriteScreen "FS", 3, 19				'Entering the SNAP program code into the Program History Filter field
	transmit

	hist_row = 8							'starting at the top of this list
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38							'reading the program history status
		If prog_hist_status = "ACTIVE" Then snap_prog_history_exists = True		'If any SPAN says'ACTIVE' there is program history
		hist_row = hist_row + 1													'going to the next row
		If hist_row = 18 Then													'going to their next page if we are at the last row
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"		'leave the list once it is at the end OR if we have found an ACTIVE SPAN
End If
If ga_status <> "ACTIVE" Then				'Looking for GA program history
	EMWriteScreen "GA", 3, 19				'Entering the GA program code into the Program History Filter field
	transmit

	hist_row = 8							'starting at the top of this list
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38							'reading the program history status
		If prog_hist_status = "ACTIVE" Then ga_prog_history_exists = True		'If any SPAN says'ACTIVE' there is program history
		hist_row = hist_row + 1													'going to the next row
		If hist_row = 18 Then													'going to their next page if we are at the last row
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"		'leave the list once it is at the end OR if we have found an ACTIVE SPAN
End If
If msa_status <> "ACTIVE" Then				'Looking for MSA program history
	EMWriteScreen "MS", 3, 19				'Entering the MSA program code into the Program History Filter field
	transmit

	hist_row = 8							'starting at the top of this list
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38							'reading the program history status
		If prog_hist_status = "ACTIVE" Then msa_prog_history_exists = True		'If any SPAN says'ACTIVE' there is program history
		hist_row = hist_row + 1													'going to the next row
		If hist_row = 18 Then													'going to their next page if we are at the last row
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"		'leave the list once it is at the end OR if we have found an ACTIVE SPAN
End If
If mfip_status <> "ACTIVE" Then				'Looking for MFIP program history
	EMWriteScreen "MF", 3, 19				'Entering the MFIP program code into the Program History Filter field
	transmit

	hist_row = 8							'starting at the top of this list
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38							'reading the program history status
		If prog_hist_status = "ACTIVE" Then mfip_prog_history_exists = True		'If any SPAN says'ACTIVE' there is program history
		hist_row = hist_row + 1													'going to the next row
		If hist_row = 18 Then													'going to their next page if we are at the last row
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"		'leave the list once it is at the end OR if we have found an ACTIVE SPAN
End If
If dwp_status <> "ACTIVE" Then
	EMWriteScreen "DW", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then dwp_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
If grh_status <> "ACTIVE" Then
	EMWriteScreen "GR", 3, 19
	transmit

	hist_row = 8
	Do
		EMReadScreen prog_hist_status, 6, hist_row, 38
		If prog_hist_status = "ACTIVE" Then grh_prog_history_exists = True
		hist_row = hist_row + 1
		If hist_row = 18 Then
			PF8
			hist_row = 8
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" then Exit Do
		End If
	Loop until prog_hist_status = "      " OR prog_hist_status = "ACTIVE"
End If
Call back_to_SELF		'reset

'saving information for error output email
script_run_lowdown = script_run_lowdown & vbCr & vbCr & "PROGRAM HISTORY:" & vbCr & "snap_prog_history_exists - " & snap_prog_history_exists & vbCr & "ga_prog_history_exists - " & ga_prog_history_exists & vbCr & "msa_prog_history_exists - " & msa_prog_history_exists & vbCr & "mfip_prog_history_exists - " & mfip_prog_history_exists & vbCr & "dwp_prog_history_exists - " & dwp_prog_history_exists & vbCr & "grh_prog_history_exists - " & grh_prog_history_exists

Call navigate_to_MAXIS_screen("STAT", "SUMM")		'Going in to STAT to read address information
EMReadScreen case_name, 22, 21, 46					'case name for address'
case_name = trim(case_name)
'Reading the information from STAT
Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three)
Call access_AREP_panel("READ", arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)
Call access_SWKR_panel("READ", swkr_name, swkr_addr_street, swkr_addr_city, swkr_addr_state, swkr_addr_zip, swkr_phone, swkr_ext, notc_to_swkr)

If arep_name <> "" Then select_a_client = select_a_client+chr(9)+"AREP - " & arep_name		'Adding AREP and SWKR to the droplist for the dialog
If swkr_name <> "" Then select_a_client = select_a_client+chr(9)+"SWKR - " & swkr_name


Do 		'BIG Loop to see if INQX is over the 9 page limit

	'MAIN DIALOG
	 Do
	 	Do
	 		err_msg = ""
			y_pos = 25

	 		Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 551, 385, "Verification of Public Assistance"
			  ButtonGroup ButtonPressed
			    Text 10, 10, 75, 10, "Requested by:"
				ComboBox 80, 5, 150, 45, select_a_client+chr(9)+verif_request_by, verif_request_by
				If snap_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "SNAP"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "SNAP Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of SNAP Needed", snap_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 200, 10, "SNAP current benefit amount appears to be $" & snap_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent SNAP Eligibility Notice appears to have been sent for benefit month: " & snap_month & "/" & snap_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, snap_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To this WCOM", snap_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", snap_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, snap_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15, snap_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", snap_view_inqx_btn
					y_pos = y_pos + 15
				End If
				If ga_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "GA"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "GA Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of GA Needed", ga_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 200, 10, "GA current benefit amount appears to be $" & ga_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent GA Eligibility Notice appears to have been sent for benefit month: " & ga_month & "/" & ga_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, ga_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To WCOM", ga_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", ga_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, ga_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15,ga_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", ga_view_inqx_btn
					y_pos = y_pos + 15
				End If
				If msa_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "MSA"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "MSA Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of MSA Needed", msa_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 200, 10, "MSA current benefit amount appears to be $" & msa_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent MSA Eligibility Notice appears to have been sent for benefit month: " & msa_month & "/" & msa_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, msa_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To WCOM", msa_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", msa_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, msa_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15,msa_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", msa_view_inqx_btn
					y_pos = y_pos + 15
				End If
				If mfip_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "MFIP"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "MFIP Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of MFIP Needed", mfip_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 350, 10, "MFIP current benefit amount appears to be; " & mfip_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent MFIP Eligibility Notice appears to have been sent for benefit month: " & mfip_month & "/" & mfip_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, mfip_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To WCOM", mfip_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", mfip_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, mfip_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15,mfip_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", mfip_view_inqx_btn
					y_pos = y_pos + 15
				End If
				If dwp_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "DWP"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "DWP Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of DWP Needed", dwp_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 200, 10, "DWP current benefit amount appears to be $" & dwp_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent DWP Eligibility Notice appears to have been sent for benefit month: " & dwp_month & "/" & dwp_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, dwp_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To this WCOM", dwp_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", dwp_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, dwp_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15, dwp_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", dwp_view_inqx_btn
					y_pos = y_pos + 15
				End If
				If grh_status = "ACTIVE" Then
					GroupBox 15, y_pos, 450, 75, "GRH"
					y_pos = y_pos + 15
					Text 20, y_pos, 120, 10, "GRH Assistance Verification to be sent via "
					DropListBox 140, y_pos - 5, 200, 45, "Select One..."+chr(9)+"Resend WCOM - Eligibility Notice"+chr(9)+"Create New MEMO with range of Months"+chr(9)+"No Verification of GRH Needed", grh_verification_method
					y_pos = y_pos + 10
					Text 25, y_pos, 200, 10, "GRH current benefit amount appears to be $" & grh_amount & "."
					y_pos = y_pos + 10
					Text 25, y_pos, 400, 10, "Most recent GRH Eligibility Notice appears to have been sent for benefit month: " & grh_month & "/" & grh_year & ". WCOM Information:"
					y_pos = y_pos + 10
					Text 30, y_pos, 200, 10, grh_wcom_text
					PushButton 225, y_pos, 100, 10, "Go To this WCOM", grh_wcom_btn
					PushButton 330, y_pos, 100, 10, "Select Different WCOM", grh_change_wcom_btn
					y_pos = y_pos + 15
					Text 25, y_pos, 105, 10, "Date range of issuance needed:"
					EditBox 130, y_pos - 5, 30, 15, grh_start_month
					Text 160, y_pos, 5, 10, "---"
					EditBox 165, y_pos - 5, 30, 15, grh_end_month
					Text 200, y_pos, 100, 10, "(use mm/yy format)"
					PushButton 330, y_pos, 100, 10, "View this INQX", grh_view_inqx_btn
					y_pos = y_pos + 15
				End If
				' If ma_status = "ACTIVE" OR msp_status = "ACTIVE" Then
				' End If
				' Text 20, 300, 200, 10, "Select the method of Notification:"
				' DropListBox 225, 295, 100, 45, "Select One..."+chr(9)+"Resend Eligibility Notices"+chr(9)+"Create new WCOM with Details", verification_method_selection
				y_pos = y_pos + 5

				If snap_status <> "ACTIVE" Then
					If snap_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "SNAP is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View SNAP Program History", snap_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of SNAP benefits issued from ", snap_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, snap_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, snap_end_month
						Text 305, y_pos, 75, 10, "(use mm/yy format)"
						PushButton 370, y_pos, 80, 10, "View this INQX", snap_view_inqx_btn
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "SNAP is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If ga_status <> "ACTIVE" Then
					If ga_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "GA is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View GA Program History", ga_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of GA benefits issued from ", ga_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, ga_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, ga_end_month
						Text 305, y_pos, 65, 10, "(use mm/yy format)"
						PushButton 370, y_pos, 80, 10, "View this INQX", ga_view_inqx_btn
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "GA is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If msa_status <> "ACTIVE" Then
					If msa_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "MSA is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View MSA Program History", msa_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of MSA benefits issued from ", msa_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, msa_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, msa_end_month
						Text 305, y_pos, 100, 10, "(use mm/yy format)"
						PushButton 370, y_pos, 80, 10, "View this INQX", msa_view_inqx_btn
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "MSA is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If mfip_status <> "ACTIVE" Then
					If mfip_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "MFIP is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View MFIP Program History", mfip_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of MFIP benefits issued from ", mfip_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, mfip_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, mfip_end_month
						Text 305, y_pos, 100, 10, "(use mm/yy format)"
						PushButton 370, y_pos, 80, 10, "View this INQX", mfip_view_inqx_btn
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "MFIP is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If dwp_status <> "ACTIVE" Then
					If dwp_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "DWP is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View DWP Program History", dwp_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of DWP benefits issued from ", dwp_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, dwp_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, dwp_end_month
						Text 305, y_pos, 100, 10, "(use mm/yy format)"
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "DWP is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If grh_status <> "ACTIVE" Then
					If grh_prog_history_exists = True Then
						Text 20, y_pos, 100, 10, "GRH is NOT currently Active"
						PushButton 120, y_pos-2, 100, 13, "View GRH Program History", grh_program_history_button
						y_pos = y_pos + 15
						CheckBox 25, y_pos, 210, 10, "Check here to include amounts of GRH benefits issued from ", grh_not_actv_memo_for_old_beneftis_checkbox
						EditBox 235, y_pos - 5, 30, 15, grh_start_month
						Text 265, y_pos, 5, 10, "---"
						EditBox 270, y_pos - 5, 30, 15, grh_end_month
						Text 305, y_pos, 100, 10, "(use mm/yy format)"
						y_pos = y_pos + 15
					Else
						Text 20, y_pos, 300, 10, "GRH is NOT currently Active and there is no ACTIVE Program history for this case."
						y_pos = y_pos + 15
					End If
				End If
				If ma_status <> "ACTIVE" AND msp_status <> "ACTIVE" Then
				End If


				OkButton 445, 365, 50, 15
				CancelButton 495, 365, 50, 15
				PushButton 35, 345, 25, 10, "CURR", CURR_button
			    PushButton 60, 345, 25, 10, "PERS", PERS_button
			    PushButton 85, 345, 25, 10, "NOTE", NOTE_button
			    PushButton 160, 345, 25, 10, "XFER", XFER_button
			    PushButton 185, 345, 25, 10, "WCOM", WCOM_button
			    PushButton 210, 345, 25, 10, "MEMO", MEMO_button
			    PushButton 35, 355, 25, 10, "PROG", PROG_button
			    PushButton 60, 355, 25, 10, "MEMB", MEMB_button
			    PushButton 85, 355, 25, 10, "REVW", REVW_button
			    PushButton 160, 355, 25, 10, "INQB", INQB_button
			    PushButton 185, 355, 25, 10, "INQD", INQD_button
			    PushButton 210, 355, 25, 10, "INQX", INQX_button
			    PushButton 35, 365, 25, 10, "SNAP", ELIG_FS_button
			    PushButton 60, 365, 25, 10, "MFIP", ELIG_MFIP_button
			    PushButton 85, 365, 25, 10, "DWP", ELIG_DWP_button
			    PushButton 110, 365, 25, 10, "GA", ELIG_GA_button
			    PushButton 135, 365, 25, 10, "MSA", ELIG_MSA_button
			    PushButton 160, 365, 25, 10, "GRH", ELIG_GRH_button
			    PushButton 185, 365, 25, 10, "HC", ELIG_HC_button
			    PushButton 210, 365, 25, 10, "SUMM", ELIG_SUMM_button
			    PushButton 235, 365, 25, 10, "DENY", ELIG_DENY_button
			  Text 250, 10, 290, 10, "NOTICE Information for Verification of Public Assistance for Case # " & MAXIS_case_number
			  ' GroupBox 5, 15, 470, 315, "Details"
			  GroupBox 5, 335, 390, 45, "Navigation"
			  Text 10, 345, 25, 10, "CASE/"
			  Text 135, 345, 25, 10, "SPEC/"
			  Text 10, 355, 25, 10, "STAT/"
			  Text 10, 365, 20, 10, "ELIG/"
			  Text 135, 355, 25, 10, "MONY/"
			EndDialog

			dialog Dialog1
			cancel_confirmation
			MAXIS_dialog_navigation
			Call leave_notice_text(False)

			If ButtonPressed > 5000 Then err_msg = "LOOP"				'these are NAV buttons - we don't want to leave the dialog if we press these OR display the err_msg
			If ButtonPressed > 1000 AND ButtonPressed < 5000 Then		'these are the WCOm search buttons - we don't want to leave the dialog if we press these OR display the err_msg
				'The program and month/year are set to generic variable based on the button pressed - because th buttons are program aligned
				If ButtonPressed = snap_change_wcom_btn Then
					selected_prog = "FS"
					notc_month = snap_month
					notc_year = snap_year
				End If
				If ButtonPressed = ga_change_wcom_btn Then
					selected_prog = "GA"
					notc_month = ga_month
					notc_year = ga_year
				End If
				If ButtonPressed = msa_change_wcom_btn Then
					selected_prog = "MS"
					notc_month = msa_month
					notc_year = msa_year
				End If
				If ButtonPressed = mfip_change_wcom_btn Then
					selected_prog = "MF"
					notc_month = mfip_month
					notc_year = mfip_year
				End If
				If ButtonPressed = dwp_change_wcom_btn Then
					selected_prog = "DW"
					notc_month = dwp_month
					notc_year = dwp_year
				End If
				If ButtonPressed = grh_change_wcom_btn Then
					selected_prog = "GR"
					notc_month = grh_month
					notc_year = grh_year
				End If

				'Here was are using the 'MEMO to WORD functionality to find and display available notices and pick one
				'    Create_List_Of_Notices(notice_panel, notices_array, selected_const, information_const, WCOM_row_const,  no_notices, specific_prog)
				Call Create_List_Of_Notices("WCOM",       notices_array, selected,       information,       WCOM_search_row, no_notices, selected_prog)

				'    Select_New_WCOM(notices_array, selected_const, information_const, WCOM_row_const,  case_number_known, allow_wcom, allow_memo, notc_month, notc_year, no_notices, specific_prog, allow_multiple_notc, allow_cancel)
				Call Select_New_WCOM(notices_array, selected,       information, 	   WCOM_search_row, True, 			   True, 	   False, 	   notc_month, notc_year, no_notices, selected_prog, False, 			  False)

				'Looking at all of the NOTICES that are in the array and applying the detail from that array into the program specific variables for this script to operate
				for each_notc = 0 to UBound(notices_array, 2)
					If notices_array(selected, each_notc) = checked Then
						If selected_prog = "FS" Then
							snap_month = notc_month
							snap_year = notc_year
							snap_wcom_text = notices_array(information, each_notc)
							snap_wcom_row = notices_array(WCOM_search_row, each_notc)
							snap_wcom_position = snap_wcom_row - 6
						End If
						If selected_prog = "GA" Then
							ga_month = notc_month
							ga_year = notc_year
							ga_wcom_text = notices_array(information, each_notc)
							ga_wcom_row = notices_array(WCOM_search_row, each_notc)
							ga_wcom_position = ga_wcom_row - 6
						End If
						If selected_prog = "MS" Then
							msa_month = notc_month
							msa_year = notc_year
							msa_wcom_text = notices_array(information, each_notc)
							msa_wcom_row = notices_array(WCOM_search_row, each_notc)
							msa_wcom_position = msa_wcom_row - 6
						End If
						If selected_prog = "MF" Then
							mfip_month = notc_month
							mfip_year = notc_year
							mfip_wcom_text = notices_array(information, each_notc)
							mfip_wcom_row = notices_array(WCOM_search_row, each_notc)
							mfip_wcom_position = mfip_wcom_row - 6
						End If
						If selected_prog = "DW" Then
							dwp_month = notc_month
							dwp_year = notc_year
							dwp_wcom_text = notices_array(information, each_notc)
							dwp_wcom_row = notices_array(WCOM_search_row, each_notc)
							dwp_wcom_position = dwp_wcom_row - 6
						End If
						If selected_prog = "GR" Then
							grh_month = notc_month
							grh_year = notc_year
							grh_wcom_text = notices_array(information, each_notc)
							grh_wcom_row = notices_array(WCOM_search_row, each_notc)
							grh_wcom_position = grh_wcom_row - 6
						End If
					End If
				next
				err_msg = "LOOP"			'do not pass go - do not collect 200 dollars'
			End If
			selected_prog = ""

			If ButtonPressed < 1000 AND ButtonPressed > 500 Then			'These are the view INQX Buttons  - we don't want to leave the dialog if we press these OR display the err_msg
				'Navigates to the desired INQX span for the specified program
				If ButtonPressed = snap_view_inqx_btn Then
					If len(snap_start_month) <> 5 OR Mid(snap_start_month, 3, 1) <> "/" OR len(snap_end_month) <> 5 OR Mid(snap_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for SNAP until you enter a start and end month in the 'mm/yy' format for SNAP."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 9, 5		'This is the SNAP place
						EMWriteScreen left(snap_start_month, 2), 6, 38
						EMWriteScreen right(snap_start_month, 2), 6, 41
						EMWriteScreen left(snap_end_month, 2), 6, 53
						EMWriteScreen right(snap_end_month, 2), 6, 56

						transmit
					End If
				End If
				If ButtonPressed = ga_view_inqx_btn   Then
					If len(ga_start_month) <> 5 OR Mid(ga_start_month, 3, 1) <> "/" OR len(ga_end_month) <> 5 OR Mid(ga_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for GA until you enter a start and end month in the 'mm/yy' format for GA."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 11, 5		'This is the GA place
						EMWriteScreen left(ga_start_month, 2), 6, 38
						EMWriteScreen right(ga_start_month, 2), 6, 41
						EMWriteScreen left(ga_end_month, 2), 6, 53
						EMWriteScreen right(ga_end_month, 2), 6, 56

						transmit
					End If
				End If
				If ButtonPressed = msa_view_inqx_btn  Then
					If len(msa_start_month) <> 5 OR Mid(msa_start_month, 3, 1) <> "/" OR len(msa_end_month) <> 5 OR Mid(msa_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for MSA until you enter a start and end month in the 'mm/yy' format for MSA."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 13, 50		'This is the MSA place
						EMWriteScreen left(msa_start_month, 2), 6, 38
						EMWriteScreen right(msa_start_month, 2), 6, 41
						EMWriteScreen left(msa_end_month, 2), 6, 53
						EMWriteScreen right(msa_end_month, 2), 6, 56

						transmit
					End If
				End If
				If ButtonPressed = mfip_view_inqx_btn Then
					If len(mfip_start_month) <> 5 OR Mid(mfip_start_month, 3, 1) <> "/" OR len(mfip_end_month) <> 5 OR Mid(mfip_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for MFIP until you enter a start and end month in the 'mm/yy' format for MFIP."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 10, 5		'This is the MFIP place
						EMWriteScreen left(mfip_start_month, 2), 6, 38
						EMWriteScreen right(mfip_start_month, 2), 6, 41
						EMWriteScreen left(mfip_end_month, 2), 6, 53
						EMWriteScreen right(mfip_end_month, 2), 6, 56

						transmit
					End If
				End If
				If ButtonPressed = dwp_view_inqx_btn Then
					If len(dwp_start_month) <> 5 OR Mid(dwp_start_month, 3, 1) <> "/" OR len(dwp_end_month) <> 5 OR Mid(dwp_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for DWP until you enter a start and end month in the 'mm/yy' format for DWP."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 17, 50		'This is the DWP place
						EMWriteScreen left(dwp_start_month, 2), 6, 38
						EMWriteScreen right(dwp_start_month, 2), 6, 41
						EMWriteScreen left(dwp_end_month, 2), 6, 53
						EMWriteScreen right(dwp_end_month, 2), 6, 56

						transmit
					End If
				End If
				If ButtonPressed = grh_view_inqx_btn  Then
					If len(grh_start_month) <> 5 OR Mid(grh_start_month, 3, 1) <> "/" OR len(grh_end_month) <> 5 OR Mid(grh_end_month, 3, 1) <> "/" Then
						MsgBox "The script cannot navigate to INQX for GRH until you enter a start and end month in the 'mm/yy' format for GRH."
					Else
						Call navigate_to_MAXIS_screen("MONY", "INQX")

						EMWriteScreen "X", 16, 50		'This is the GRH place
						EMWriteScreen left(grh_start_month, 2), 6, 38
						EMWriteScreen right(grh_start_month, 2), 6, 41
						EMWriteScreen left(grh_end_month, 2), 6, 53
						EMWriteScreen right(grh_end_month, 2), 6, 56

						transmit
					End If
				End If
				err_msg = "LOOP"			'do not pass go - do not collect 200 dollars

			End If
			If ButtonPressed < 500 AND ButtonPressed > 100 Then				'These are the open WCOM Buttons - we don't want to leave the dialog if we press these OR display the err_msg'
				'setting hte program specific WCOM information to generic variables
				If ButtonPressed = snap_wcom_btn Then
					wcom_row_to_open = snap_wcom_row
					wcom_month = snap_month
					wcom_year = snap_year
				End If
				If ButtonPressed = ga_wcom_btn Then
					wcom_row_to_open = ga_wcom_row
					wcom_month = ga_month
					wcom_year = ga_year
				End If
				If ButtonPressed = msa_wcom_btn Then
					wcom_row_to_open = msa_wcom_row
					wcom_month = msa_month
					wcom_year = msa_year
				End If
				If ButtonPressed = mfip_wcom_btn Then
					wcom_row_to_open = mfip_wcom_row
					wcom_month = mfip_month
					wcom_year = mfip_year
				End If
				If ButtonPressed = dwp_wcom_btn Then
					wcom_row_to_open = dwp_wcom_row
					wcom_month = dwp_month
					wcom_year = dwp_year
				End If
				If ButtonPressed = grh_wcom_btn Then
					wcom_row_to_open = grh_wcom_row
					wcom_month = grh_month
					wcom_year = grh_year
				End If

				'Navigate to the correct WCOM
				Call navigate_to_MAXIS_screen("SPEC", "WCOM")
				EMWriteScreen wcom_month, 3, 46
				EMWriteScreen wcom_year, 3, 51
				transmit
				EMWriteScreen "X", wcom_row_to_open, 13
				'Asks if they want the script to actually OPEN the WCOM
				open_wcom = MsgBox("The WCOM Notice has been selected." & vbCr & vbCr & "Would you like to open the notice?", vbQuestion + vbYesNo, "WCOM selected")
				If open_wcom = vbYes Then
					transmit
				Else
					EMWriteScreen " ", wcom_row_to_open, 13
				End If

				err_msg = "LOOP"			'do not pass go - do not collect 200 dollars
			End If

			If ButtonPressed > 50 AND ButtonPressed < 100 Then			'These are the 'view Program History' buttons - we don't want to leave the dialog if we press these OR display the err_msg
				If ButtonPressed = snap_program_history_button Then prog_to_search = "FS"
				If ButtonPressed = ga_program_history_button Then prog_to_search = "GA"
				If ButtonPressed = msa_program_history_button Then prog_to_search = "MS"
				If ButtonPressed = mfip_program_history_button Then prog_to_search = "MF"
				If ButtonPressed = dwp_program_history_button Then prog_to_search = "DW"
				If ButtonPressed = grh_program_history_button Then prog_to_search = "GR"
				If ButtonPressed = hc_program_history_button Then
					'WAY MORE STUFF GOES HERE
				End If

				'Opening PPROGRAM History
				Call navigate_to_MAXIS_screen("CASE", "CURR")
				EMWriteScreen "X", 4, 9
				transmit
				EMWriteScreen prog_to_search, 3, 19
				transmit

				err_msg = "LOOP"			'do not pass go - do not collect 200 dollars
			End If

			If err_msg <> "LOOP" Then			'TRYING TO PASS GO AND COLLECT 200 DOLLARS
				snap_start_month = trim(snap_start_month)
				snap_end_month = trim(snap_end_month)
				ga_start_month = trim(ga_start_month)
				ga_end_month = trim(ga_end_month)
				msa_start_month = trim(msa_start_month)
				msa_end_month = trim(msa_end_month)
				mfip_start_month = trim(mfip_start_month)
				mfip_end_month = trim(mfip_end_month)
				dwp_start_month = trim(dwp_start_month)
				dwp_end_month = trim(dwp_end_month)
				grh_start_month = trim(grh_start_month)
				grh_end_month = trim(grh_end_month)
				verif_request_by = trim(verif_request_by)

				'ERROR HANDLING to be sure the correct details have been included.
				If verif_request_by = "" or verif_request_by = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Indicate who is requesting the information. You can select someone from the household or write in the name of the person. Please only provide information to individuals who have the right to access the information."
				If snap_status = "ACTIVE" Then
					If snap_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since SNAP is active, indicate if Verification of SNAP benefits is needed, and if so, which method works best."
					If snap_verification_method = "Resend WCOM - Eligibility Notice" AND snap_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of SNAP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If snap_verification_method = "Create New MEMO with range of Months" Then
					 	If len(snap_start_month) <> 5 OR Mid(snap_start_month, 3, 1) <> "/" OR len(snap_end_month) <> 5 OR Mid(snap_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of SNAP issuance history to be sent as verification of Active SNAP, enter a start and end month in the 'mm/yy' format."
						If len(snap_end_month) = 5 AND Mid(snap_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(snap_end_month, 2) & "/1/" & right(snap_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The SNAP end month of " & snap_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								snap_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If
				If ga_status = "ACTIVE" Then
					If ga_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since GA is active, indicate if Verification of GA benefits is needed, and if so, which method works best."
					If ga_verification_method = "Resend WCOM - Eligibility Notice" AND ga_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of GA, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If ga_verification_method = "Create New MEMO with range of Months" Then
					 	If len(ga_start_month) <> 5 OR Mid(ga_start_month, 3, 1) <> "/" OR len(ga_end_month) <> 5 OR Mid(ga_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GA issuance history to be sent as verification of Active GA, enter a start and end month in the 'mm/yy' format."
						If len(ga_end_month) = 5 AND Mid(ga_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(ga_end_month, 2) & "/1/" & right(ga_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The GA end month of " & ga_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								ga_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If
				If msa_status = "ACTIVE" Then
					If msa_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since MSA is active, indicate if Verification of MSA benefits is needed, and if so, which method works best."
					If msa_verification_method = "Resend WCOM - Eligibility Notice" AND msa_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of MSA, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If msa_verification_method = "Create New MEMO with range of Months" Then
					 	If len(msa_start_month) <> 5 OR Mid(msa_start_month, 3, 1) <> "/" OR len(msa_end_month) <> 5 OR Mid(msa_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MSA issuance history to be sent as verification of Active MSA, enter a start and end month in the 'mm/yy' format."
						If len(msa_end_month) = 5 AND Mid(msa_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(msa_end_month, 2) & "/1/" & right(msa_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The MSA end month of " & msa_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								msa_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If
				If mfip_status = "ACTIVE" Then
					If mfip_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since MFIP is active, indicate if Verification of MFIP benefits is needed, and if so, which method works best."
					If mfip_verification_method = "Resend WCOM - Eligibility Notice" AND mfip_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of MFIP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If mfip_verification_method = "Create New MEMO with range of Months" Then
					 	If len(mfip_start_month) <> 5 OR Mid(mfip_start_month, 3, 1) <> "/" OR len(mfip_end_month) <> 5 OR Mid(mfip_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MFIP issuance history to be sent as verification of Active MFIP, enter a start and end month in the 'mm/yy' format."
						If len(mfip_end_month) = 5 AND Mid(mfip_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(mfip_end_month, 2) & "/1/" & right(mfip_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The MFIP end month of " & mfip_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								mfip_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If
				If dwp_status = "ACTIVE" Then
					If dwp_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since DWP is active, indicate if Verification of DWP benefits is needed, and if so, which method works best."
					If dwp_verification_method = "Resend WCOM - Eligibility Notice" AND dwp_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of DWP, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If dwp_verification_method = "Create New MEMO with range of Months" Then
						If len(dwp_start_month) <> 5 OR Mid(dwp_start_month, 3, 1) <> "/" OR len(dwp_end_month) <> 5 OR Mid(dwp_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of DWP issuance history to be sent as verification of Active DWP, enter a start and end month in the 'mm/yy' format."
						If len(dwp_end_month) = 5 AND Mid(dwp_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(dwp_end_month, 2) & "/1/" & right(dwp_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The DWP end month of " & dwp_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								dwp_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If
				If grh_status = "ACTIVE" Then
				 	If grh_verification_method = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since GRH is active, indicate if Verification of GRH benefits is needed, and if so, which method works best."
					If grh_verification_method = "Resend WCOM - Eligibility Notice" AND grh_wcom_text = "NO WCOM Found" then err_msg = err_msg & vbNewLine & "* Since you are selecting a WCOM to be resent as verification of GRH, use the 'Select Different WCOM' button to select the correct WCOM since none was found."
					If grh_verification_method = "Create New MEMO with range of Months" Then
					 	If len(grh_start_month) <> 5 OR Mid(grh_start_month, 3, 1) <> "/" OR len(grh_end_month) <> 5 OR Mid(grh_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GRH issuance history to be sent as verification of Active GRH, enter a start and end month in the 'mm/yy' format."
						If len(grh_end_month) = 5 AND Mid(grh_end_month, 3, 1) = "/" Then
							first_day_of_end_month = left(grh_end_month, 2) & "/1/" & right(grh_end_month, 2)
							first_day_of_end_month = DateAdd("d", 0, first_day_of_end_month)
							If DateDiff("d", date, first_day_of_end_month) > 0 Then
								err_msg = err_msg & vbNewLine & "* We should not send information about benefits issued for a future month. The GRH end month of " & grh_end_month & " has been changed to " & CM_mo & "/" & CM_yr & " as benefits have not been issued for a future month and would not provide good information to the resident."
								grh_end_month = CM_mo & "/" & CM_yr
							End If
						End If
					End If
				End If

				If snap_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(snap_start_month) <> 5 OR Mid(snap_start_month, 3, 1) <> "/" OR len(snap_end_month) <> 5 OR Mid(snap_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of SNAP issuance history to be sent as verification of Previous SNAP Eligibility, enter a start and end month in the 'mm/yy' format."
				End If
				If ga_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(ga_start_month) <> 5 OR Mid(ga_start_month, 3, 1) <> "/" OR len(ga_end_month) <> 5 OR Mid(ga_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GA issuance history to be sent as verification of Previous GA Eligibility, enter a start and end month in the 'mm/yy' format."
				End If
				If msa_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(msa_start_month) <> 5 OR Mid(msa_start_month, 3, 1) <> "/" OR len(msa_end_month) <> 5 OR Mid(msa_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MSA issuance history to be sent as verification of Previous MSA Eligibility, enter a start and end month in the 'mm/yy' format."
				End If
				If mfip_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(mfip_start_month) <> 5 OR Mid(mfip_start_month, 3, 1) <> "/" OR len(mfip_end_month) <> 5 OR Mid(mfip_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of MFIP issuance history to be sent as verification of Previous MFIP Eligibility, enter a start and end month in the 'mm/yy' format."
				End If
				If dwp_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(dwp_start_month) <> 5 OR Mid(dwp_start_month, 3, 1) <> "/" OR len(dwp_end_month) <> 5 OR Mid(dwp_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of DWP issuance history to be sent as verification of Previous DWP Eligibility, enter a start and end month in the 'mm/yy' format."
				End If
				If grh_not_actv_memo_for_old_beneftis_checkbox = checked Then
					If len(grh_start_month) <> 5 OR Mid(grh_start_month, 3, 1) <> "/" OR len(grh_end_month) <> 5 OR Mid(grh_end_month, 3, 1) <> "/" Then err_msg = err_msg & vbNewLine & "* Since you are creating a MEMO of GRH issuance history to be sent as verification of Previous GRH Eligibility, enter a start and end month in the 'mm/yy' format."
				End If

				'Displaying the Error handling
				If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
			End If
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False

	'saving information for error output email
	script_run_lowdown = script_run_lowdown & vbCr & vbCr & "NOTICE Selections:"
	script_run_lowdown = script_run_lowdown & vbCr & "SNAP - " & vbCr & "snap_verification_method - " & snap_verification_method & vbCr & "SNAP months: " & snap_start_month & " to " & snap_end_month & vbCr & "WCOM text - " & snap_wcom_text & vbCr & "Closed Program checkbox - " & snap_not_actv_memo_for_old_beneftis_checkbox
	script_run_lowdown = script_run_lowdown & vbCr & "GA - " & vbCr & "ga_verification_method - " & ga_verification_method & vbCr & "GA months: " & ga_start_month & " to " & ga_end_month & vbCr & "WCOM text - " & ga_wcom_text & vbCr & "Closed Program checkbox - " & ga_not_actv_memo_for_old_beneftis_checkbox
	script_run_lowdown = script_run_lowdown & vbCr & "MSA - " & vbCr & "msa_verification_method - " & msa_verification_method & vbCr & "MSA months: " & msa_start_month & " to " & msa_end_month & vbCr & "WCOM text - " & msa_wcom_text & vbCr & "Closed Program checkbox - " & msa_not_actv_memo_for_old_beneftis_checkbox
	script_run_lowdown = script_run_lowdown & vbCr & "MFIP - " & vbCr & "mfip_verification_method - " & mfip_verification_method & vbCr & "MFIP months: " & mfip_start_month & " to " & mfip_end_month & vbCr & "WCOM text - " & mfip_wcom_text & vbCr & "Closed Program checkbox - " & mfip_not_actv_memo_for_old_beneftis_checkbox
	script_run_lowdown = script_run_lowdown & vbCr & "GRH - " & vbCr & "grh_verification_method - " & grh_verification_method & vbCr & "GRH months: " & grh_start_month & " to " & grh_end_month & vbCr & "WCOM text - " & grh_wcom_text & vbCr & "Closed Program checkbox - " & grh_not_actv_memo_for_old_beneftis_checkbox

	'Setting what kind of notices are needed
	If snap_not_actv_memo_for_old_beneftis_checkbox = checked Then snap_verification_method = "Create New MEMO with range of Months"
	If ga_not_actv_memo_for_old_beneftis_checkbox = checked Then ga_verification_method = "Create New MEMO with range of Months"
	If msa_not_actv_memo_for_old_beneftis_checkbox = checked Then msa_verification_method = "Create New MEMO with range of Months"
	If mfip_not_actv_memo_for_old_beneftis_checkbox = checked Then mfip_verification_method = "Create New MEMO with range of Months"
	If dwp_not_actv_memo_for_old_beneftis_checkbox = checked Then dwp_verification_method = "Create New MEMO with range of Months"
	If grh_not_actv_memo_for_old_beneftis_checkbox = checked Then grh_verification_method = "Create New MEMO with range of Months"

	create_memo = False
	If snap_verification_method = "Create New MEMO with range of Months" Then create_memo = True
	If ga_verification_method = "Create New MEMO with range of Months" Then create_memo = True
	If msa_verification_method = "Create New MEMO with range of Months" Then create_memo = True
	If mfip_verification_method = "Create New MEMO with range of Months" Then create_memo = True
	If dwp_verification_method = "Create New MEMO with range of Months" Then create_memo = True
	If grh_verification_method = "Create New MEMO with range of Months" Then create_memo = True

	resend_wcom = False
	If snap_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
	If ga_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
	If msa_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
	If mfip_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
	If dwp_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True
	If grh_verification_method = "Resend WCOM - Eligibility Notice" Then resend_wcom = True

	previous_active_prog_memo = False
	If snap_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True
	If ga_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True
	If msa_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True
	If mfip_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True
	If dwp_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True
	If grh_not_actv_memo_for_old_beneftis_checkbox = checked then previous_active_prog_memo = True

	'If no kind of nitice has been requested - the script will End.
	If create_memo = False AND resend_wcom = False AND previous_active_prog_memo = False Then
		end_msg = "No NOTICE SENT"& vbCr & vbCr & "No notices were requested for any program and there is no additional action for the script to take or actions to note." & vbCr & vbCr & "This does not mean there was an error. If you intended to select a MEMO or WCOM for one of the programs, rerun the script and enter the selections for the appropriate notice on the correct program."
		script_end_procedure_with_error_report(end_msg)
	End If

	call back_to_SELF

	too_many_SNAP_INQX_pages = False
	too_many_MFIP_INQX_pages = False
	too_many_GA_INQX_pages = False
	too_many_MSA_INQX_pages = False
	too_many_DWP_INQX_pages = False
	too_many_GRH_INQX_pages = False

	If create_memo = True Then		'If there are any MEMOs needed we need to read INQX for all the specified programs and dates and create arrays of the benefit months for each program
		If snap_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			SNAP_total = 0
			SNAP_MEMO_rows_needed = 2

			first_date_of_range = replace(snap_start_month, "/", "/01/")			'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			last_date_of_range = replace(snap_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			SNAP_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				SNAP_expected_dates_array = SNAP_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(SNAP_expected_dates_array, "~") = 0 Then
				SNAP_expected_dates_array = Array(SNAP_expected_dates_array)
			Else
				SNAP_expected_dates_array = split(SNAP_expected_dates_array, "~")
			End If

			EMWriteScreen "X", 9, 5		'This is the SNAP place						'Opening the right detail in INQX based on the dates and program
			EMWriteScreen left(snap_start_month, 2), 6, 38
			EMWriteScreen right(snap_start_month, 2), 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit

			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance) = SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance) + tran_amount
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve SNAP_ISSUANCE_ARRAY(last_const, msg_counter)
							SNAP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = tran_amount
							SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6
					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_SNAP_INQX_pages = True
					If too_many_SNAP_INQX_pages = True Then
						ReDim SNAP_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			SNAP_dates_array = ""			'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_amount = SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance)
				total_amount = total_amount & ""
				If InStr(total_amount, ".") = 0 Then
					total_amount = left(total_amount & ".00        ", 8)
				Else
					total_amount = left(total_amount & "        ", 8)
				End If
				SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
				SNAP_dates_array = SNAP_dates_array & "~" & SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in SNAP_expected_dates_array					'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)		'Look at all the found months - if they match - indicate that here
					If DateDiff("d", SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve SNAP_ISSUANCE_ARRAY(last_const, msg_counter)
					SNAP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = 0
					SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ 0.00     issued for " & SNAP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
					SNAP_dates_array = SNAP_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(SNAP_dates_array, 1) = "~" Then SNAP_dates_array = right(SNAP_dates_array, len(SNAP_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(SNAP_dates_array, "~") = 0 Then
				SNAP_dates_array = Array(SNAP_dates_array)
			Else
				SNAP_dates_array = split(SNAP_dates_array, "~")
			End If
			Call sort_dates(SNAP_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in SNAP_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						snap_msg_display = snap_msg_display & vbCr & SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						SNAP_total = SNAP_total + SNAP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance)
						SNAP_MEMO_rows_needed = SNAP_MEMO_rows_needed + 1
					End If
				Next
			Next

			' MsgBox "SNAP - This is the list" & snap_msg_display & vbCr & "TOTAL SNAP: $" & SNAP_total
			PF3
		End If

		If ga_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			GA_total = 0

			first_date_of_range = replace(ga_start_month, "/", "/01/")			'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			last_date_of_range = replace(ga_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			GA_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				GA_expected_dates_array = GA_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(GA_expected_dates_array, "~") = 0 Then
				GA_expected_dates_array = Array(GA_expected_dates_array)
			Else
				GA_expected_dates_array = split(GA_expected_dates_array, "~")
			End If

			EMWriteScreen "X", 11, 5		'This is the GA place					'Opening the right detail in INQX based on the dates and program
			EMWriteScreen left(ga_start_month, 2), 6, 38
			EMWriteScreen right(ga_start_month, 2), 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit

			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(GA_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = GA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								GA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) = GA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) + tran_amount
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve GA_ISSUANCE_ARRAY(last_const, msg_counter)
							GA_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							GA_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = tran_amount
							GA_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6

					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_GA_INQX_pages = True
					If too_many_GA_INQX_pages = True Then
						ReDim GA_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			GA_dates_array = ""				'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(GA_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_amount = GA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
				total_amount = total_amount & ""
				If InStr(total_amount, ".") = 0 Then
					total_amount = left(total_amount & ".00        ", 8)
				Else
					total_amount = left(total_amount & "        ", 8)
				End If
				GA_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & GA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
				GA_dates_array = GA_dates_array & "~" & GA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in GA_expected_dates_array						'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(GA_ISSUANCE_ARRAY, 2)			'Look at all the found months - if they match - indicate that here
					If DateDiff("d", GA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve GA_ISSUANCE_ARRAY(last_const, msg_counter)
					GA_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					GA_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = 0
					GA_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					GA_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ 0.00     issued for " & GA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
					GA_dates_array = GA_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(GA_dates_array, 1) = "~" Then GA_dates_array = right(GA_dates_array, len(GA_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(GA_dates_array, "~") = 0 Then
				GA_dates_array = Array(GA_dates_array)
			Else
				GA_dates_array = split(GA_dates_array, "~")
			End If
			Call sort_dates(GA_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in GA_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(GA_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, GA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						ga_msg_display = ga_msg_display & vbCr & GA_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						GA_total = GA_total + GA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
					End If
				Next
			Next

			' MsgBox "GA - This is the list" & ga_msg_display & vbCr & "TOTAL GA: $" & GA_total
			PF3
		End If

		If msa_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			MSA_total = 0

			first_date_of_range = replace(msa_start_month, "/", "/01/")				'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			last_date_of_range = replace(msa_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			MSA_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				MSA_expected_dates_array = MSA_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(MSA_expected_dates_array, "~") = 0 Then
				MSA_expected_dates_array = Array(MSA_expected_dates_array)
			Else
				MSA_expected_dates_array = split(MSA_expected_dates_array, "~")
			End If

			EMWriteScreen "X", 13, 50		'This is the MSA place					'Opening the right detail in INQX based on the dates and program
			EMWriteScreen left(msa_start_month, 2), 6, 38
			EMWriteScreen right(msa_start_month, 2), 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit
			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(MSA_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = MSA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								MSA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) = MSA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) + tran_amount
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve MSA_ISSUANCE_ARRAY(last_const, msg_counter)
							MSA_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							MSA_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = tran_amount
							MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6

					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_MSA_INQX_pages = True
					If too_many_MSA_INQX_pages = True Then
						ReDim MSA_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			MSA_dates_array = ""			'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(MSA_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_amount = MSA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
				total_amount = total_amount & ""
				If InStr(total_amount, ".") = 0 Then
					total_amount = left(total_amount & ".00        ", 8)
				Else
					total_amount = left(total_amount & "        ", 8)
				End If
				MSA_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & MSA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
				MSA_dates_array = MSA_dates_array & "~" & MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in MSA_expected_dates_array					'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(MSA_ISSUANCE_ARRAY, 2)		'Look at all the found months - if they match - indicate that here
					If DateDiff("d", MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve MSA_ISSUANCE_ARRAY(last_const, msg_counter)
					MSA_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					MSA_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = 0
					MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					MSA_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ 0.00     issued for " & MSA_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
					MSA_dates_array = MSA_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(MSA_dates_array, 1) = "~" Then MSA_dates_array = right(MSA_dates_array, len(MSA_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(MSA_dates_array, "~") = 0 Then
				MSA_dates_array = Array(MSA_dates_array)
			Else
				MSA_dates_array = split(MSA_dates_array, "~")
			End If
			Call sort_dates(MSA_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in MSA_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(MSA_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						msa_msg_display = msa_msg_display & vbCr & MSA_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						MSA_total = MSA_total + MSA_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
					End If
				Next
			Next

			' MsgBox "MSA - This is the list" & msa_msg_display & vbCr & "MSA Total: $" & MSA_total
			PF3
		End If

		If mfip_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			MFIP_Cash_total = 0
			MFIP_Food_total = 0

			first_date_of_range = replace(mfip_start_month, "/", "/01/")			'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			mfip_search_month = DateAdd("m", -1, first_date_of_range)				'MFIP issues the day before the benefit month, so we need to search starting a month earlier
			search_month = right("00" & DatePart("m", mfip_search_month), 2)
			search_year = right(DatePart("yyyy", mfip_search_month), 2)
			last_date_of_range = replace(mfip_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			MFIP_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				MFIP_expected_dates_array = MFIP_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(MFIP_expected_dates_array, "~") = 0 Then
				MFIP_expected_dates_array = Array(MFIP_expected_dates_array)
			Else
				MFIP_expected_dates_array = split(MFIP_expected_dates_array, "~")
			End If

			EMWriteScreen "X", 10, 5		'This is the MFIP place					'Opening the right detail in INQX based on the dates and program
			EMWriteScreen search_month, 6, 38
			EMWriteScreen search_year, 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit

			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen ben_type, 2, inqx_row, 19
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(MFIP_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = MFIP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								If ben_type = "FS" Then
									MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance) = MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance) + tran_amount
								End If
								If ben_type = "MF" OR ben_type = "HG" Then
									MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) = MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance) + tran_amount
								End If
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve MFIP_ISSUANCE_ARRAY(last_const, msg_counter)
							MFIP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							MFIP_ISSUANCE_ARRAY(grant_amount_const, msg_counter) = tran_amount
							If ben_type = "FS" Then
								MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = tran_amount
								MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = 0
							End If
							If ben_type = "MF" OR ben_type = "HG" Then
								MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = 0
								MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = tran_amount
							End If
							MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6
					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_MFIP_INQX_pages = True
					If too_many_MFIP_INQX_pages = True Then
						ReDim MFIP_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			MFIP_dates_array = ""			'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(MFIP_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_cash_amount = MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
				total_cash_amount = total_cash_amount & ""
				If InStr(total_cash_amount, ".") = 0 Then
					total_cash_amount = left(total_cash_amount & ".00        ", 8)
				Else
					total_cash_amount = left(total_cash_amount & "        ", 8)
				End If

				total_snap_amount = MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance)
				total_snap_amount = total_snap_amount & ""
				If InStr(total_snap_amount, ".") = 0 Then
					total_snap_amount = left(total_snap_amount & ".00        ", 8)
				Else
					total_snap_amount = left(total_snap_amount & "        ", 8)
				End If

				MFIP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = MFIP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) & " - CASH: $ " & total_cash_amount & " and FOOD: $ " & total_snap_amount
				MFIP_dates_array = MFIP_dates_array & "~" & MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in MFIP_expected_dates_array					'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(MFIP_ISSUANCE_ARRAY, 2)		'Look at all the found months - if they match - indicate that here
					If DateDiff("d", MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve MFIP_ISSUANCE_ARRAY(last_const, msg_counter)
					MFIP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = 0
					MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, msg_counter) = 0
					MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					MFIP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = MFIP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) & " - CASH: $ 0.00     and FOOD: $ 0.00    "
					MFIP_dates_array = MFIP_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(MFIP_dates_array, 1) = "~" Then MFIP_dates_array = right(MFIP_dates_array, len(MFIP_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(MFIP_dates_array, "~") = 0 Then
				MFIP_dates_array = Array(MFIP_dates_array)
			Else
				MFIP_dates_array = split(MFIP_dates_array, "~")
			End If
			Call sort_dates(MFIP_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in MFIP_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(MFIP_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						mfip_msg_display = mfip_msg_display & vbCr & MFIP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						MFIP_Cash_total = MFIP_Cash_total + MFIP_ISSUANCE_ARRAY(cash_grant_amount_const, each_known_issuance)
						MFIP_Food_total = MFIP_Food_total + MFIP_ISSUANCE_ARRAY(snap_grant_amount_const, each_known_issuance)
					End If
				Next
			Next

			' MsgBox "MFIP - This is the list" & mfip_msg_display & vbCr & "MFIP Cash Total: $" & MFIP_Cash_total & vbCr & "MFIP Food Total: $" & MFIP_Food_total
			PF3
		End If

		If dwp_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			DWP_total = 0
			DWP_MEMO_rows_needed = 2

			first_date_of_range = replace(dwp_start_month, "/", "/01/")			'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			last_date_of_range = replace(dwp_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			DWP_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				DWP_expected_dates_array = DWP_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(DWP_expected_dates_array, "~") = 0 Then
				DWP_expected_dates_array = Array(DWP_expected_dates_array)
			Else
				DWP_expected_dates_array = split(DWP_expected_dates_array, "~")
			End If

			EMWriteScreen "X", 17, 50		'This is the DWP place						'Opening the right detail in INQX based on the dates and program
			EMWriteScreen left(dwp_start_month, 2), 6, 38
			EMWriteScreen right(dwp_start_month, 2), 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit

			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(DWP_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = DWP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, each_known_issuance) = DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, each_known_issuance) + tran_amount
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve DWP_ISSUANCE_ARRAY(last_const, msg_counter)
							DWP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, msg_counter) = tran_amount
							DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6
					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_DWP_INQX_pages = True
					If too_many_DWP_INQX_pages = True Then
						ReDim DWP_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			DWP_dates_array = ""			'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(DWP_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_amount = DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, each_known_issuance)
				total_amount = total_amount & ""
				If InStr(total_amount, ".") = 0 Then
					total_amount = left(total_amount & ".00        ", 8)
				Else
					total_amount = left(total_amount & "        ", 8)
				End If
				DWP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & DWP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
				DWP_dates_array = DWP_dates_array & "~" & DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in DWP_expected_dates_array					'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(DWP_ISSUANCE_ARRAY, 2)		'Look at all the found months - if they match - indicate that here
					If DateDiff("d", DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve DWP_ISSUANCE_ARRAY(last_const, msg_counter)
					DWP_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, msg_counter) = 0
					DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					DWP_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ 0.00     issued for " & DWP_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
					DWP_dates_array = DWP_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(DWP_dates_array, 1) = "~" Then DWP_dates_array = right(DWP_dates_array, len(DWP_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(DWP_dates_array, "~") = 0 Then
				DWP_dates_array = Array(DWP_dates_array)
			Else
				DWP_dates_array = split(DWP_dates_array, "~")
			End If
			Call sort_dates(DWP_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in DWP_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(DWP_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						dwp_msg_display = dwp_msg_display & vbCr & DWP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						DWP_total = DWP_total + DWP_ISSUANCE_ARRAY(dwp_grant_amount_const, each_known_issuance)
						DWP_MEMO_rows_needed = DWP_MEMO_rows_needed + 1
					End If
				Next
			Next

			' MsgBox "DWP - This is the list" & dwp_msg_display & vbCr & "TOTAL DWP: $" & DWP_total
			PF3
		End If

		If grh_verification_method = "Create New MEMO with range of Months" Then
			Call navigate_to_MAXIS_screen("MONY", "INQX")							'Go to where the benefit amounts are listed

			GRH_total = 0
			GRH_MEMO_rows_needed = 2

			first_date_of_range = replace(grh_start_month, "/", "/01/")			'setting the month for start and end dates as actual dates
			first_date_of_range = DateAdd("d", 0, first_date_of_range)
			last_date_of_range = replace(grh_end_month, "/", "/01/")
			last_date_of_range = DateAdd("d", 0, last_date_of_range)

			GRH_expected_dates_array = first_date_of_range							'creating an array of all of the months in the range
			each_date = first_date_of_range
			Do
				each_date = DateAdd("m", 1, each_date)
				GRH_expected_dates_array = GRH_expected_dates_array & "~" & each_date
			Loop until each_date = last_date_of_range

			If InStr(GRH_expected_dates_array, "~") = 0 Then
				GRH_expected_dates_array = Array(GRH_expected_dates_array)
			Else
				GRH_expected_dates_array = split(GRH_expected_dates_array, "~")
			End If


			EMWriteScreen "X", 16, 50		'This is the GRH place
			EMWriteScreen left(grh_start_month, 2), 6, 38
			EMWriteScreen right(grh_start_month, 2), 6, 41
			EMWriteScreen CM_plus_1_mo, 6, 53
			EMWriteScreen CM_plus_1_yr, 6, 56

			transmit

			inqx_row = 6															'Read all of the information on INQX
			msg_counter = 0
			Do
				EMReadScreen issued_date, 8, inqx_row, 7
				EMReadScreen tran_amount, 8, inqx_row, 38
				EMReadScreen from_month, 2, inqx_row, 62
				EMReadScreen from_year, 2, inqx_row, 68

				issued_date = trim(issued_date)
				tran_amount = trim(tran_amount)

				If issued_date <> "" Then
					from_date = from_month & "/1/" & from_year						'making the date a date and making it the 1st of the month (this accounts for proration)
					from_date = DateAdd("d", 0, from_date)
					'Only accept if the date is equal to or after the first date and equal to or before the last date
					If DateDiff("d", from_date, first_date_of_range) <= 0 AND DateDiff("d", from_date, last_date_of_range) >= 0 Then

						benefit_month = from_month & "/" & from_year
						tran_amount = tran_amount * 1								'this must be a NUMBER
						ammount_added_in = False
						For each_known_issuance = 0 to UBound(GRH_ISSUANCE_ARRAY, 2)		'reading to see if the benefit month is already in the array so we can combine the benefit amounts
							If benefit_month = GRH_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance) Then
								GRH_ISSUANCE_ARRAY(grh_grant_amount_const, each_known_issuance) = GRH_ISSUANCE_ARRAY(grh_grant_amount_const, each_known_issuance) + tran_amount
								ammount_added_in = True
							End If
						Next
						If ammount_added_in = False Then							'if the benefit month was NOT found - create a new array instance for that benefit month.
							ReDim Preserve GRH_ISSUANCE_ARRAY(last_const, msg_counter)
							GRH_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = benefit_month
							GRH_ISSUANCE_ARRAY(grh_grant_amount_const, msg_counter) = tran_amount
							GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = from_date
							msg_counter = msg_counter + 1
						End If
					End If
				End If

				inqx_row = inqx_row + 1		'go to the next line/page
				If inqx_row = 18 Then
					PF8
					inqx_row = 6

					EMReadScreen more_thanb_9_pages_msg, 38, 24, 2
					If more_thanb_9_pages_msg = "CAN NOT PAGE THROUGH MORE THAN 9 PAGES" Then too_many_GRH_INQX_pages = True
					If too_many_GRH_INQX_pages = True Then
						ReDim GRH_ISSUANCE_ARRAY(last_const, 0)
						Exit Do
					End if
					EMreadScreen end_of_list, 9, 24, 14
					if end_of_list = "LAST PAGE" Then Exit Do
				End If
			Loop until issued_date = ""		'go until the end of the list
			GRH_dates_array = ""			'we need an array of the dates ONLY
			For each_known_issuance = 0 to UBound(GRH_ISSUANCE_ARRAY, 2)			'Now we loop through all of the found benefit months and create the formatting for the MEMO
				total_amount = GRH_ISSUANCE_ARRAY(grh_grant_amount_const, each_known_issuance)
				total_amount = total_amount & ""
				If InStr(total_amount, ".") = 0 Then
					total_amount = left(total_amount & ".00        ", 8)
				Else
					total_amount = left(total_amount & "        ", 8)
				End If
				GRH_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ " & total_amount & " issued for " & GRH_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
				GRH_dates_array = GRH_dates_array & "~" & GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)		'adding to the array of all the dates
			Next
			For each expected_month in GRH_expected_dates_array					'Now we loop through ALL the months we expected to find in the range - this is so we can add $0 issuance months as 0
				issuance_found = False
				For each_known_issuance = 0 to UBound(GRH_ISSUANCE_ARRAY, 2)		'Look at all the found months - if they match - indicate that here
					If DateDiff("d", GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance), expected_month) = 0 Then issuance_found = True
				Next
				If issuance_found = False Then										'If no month was found - add another array instance with a $0 benefit amount listed
					ReDim Preserve GRH_ISSUANCE_ARRAY(last_const, msg_counter)
					GRH_ISSUANCE_ARRAY(benefit_month_const, msg_counter) = right("00" & DatePart("m", expected_month), 2) & "/" & right(DatePart("yyyy", expected_month), 2)
					GRH_ISSUANCE_ARRAY(snap_grant_amount_const, msg_counter) = 0
					GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, msg_counter) = expected_month
					GRH_ISSUANCE_ARRAY(note_message_const, each_known_issuance) = "$ 0.00     issued for " & GRH_ISSUANCE_ARRAY(benefit_month_const, each_known_issuance)
					GRH_dates_array = GRH_dates_array & "~" & expected_month
					msg_counter = msg_counter + 1
				End If
			Next
			If left(GRH_dates_array, 1) = "~" Then GRH_dates_array = right(GRH_dates_array, len(GRH_dates_array) - 1)		'creating an array of all of the 'from dates'
			If Instr(GRH_dates_array, "~") = 0 Then
				GRH_dates_array = Array(GRH_dates_array)
			Else
				GRH_dates_array = split(GRH_dates_array, "~")
			End If
			Call sort_dates(GRH_dates_array)		'This function takes all the dates in an array and put them in order from oldest to newest

			for each ordered_date in GRH_dates_array		'Now doing some counting and totalling
				For each_known_issuance = 0 to UBound(GRH_ISSUANCE_ARRAY, 2)
					If DateDiff("d", ordered_date, GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
						grh_msg_display = grh_msg_display & vbCr & GRH_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
						GRH_total = GRH_total + GRH_ISSUANCE_ARRAY(grh_grant_amount_const, each_known_issuance)
						GRH_MEMO_rows_needed = GRH_MEMO_rows_needed + 1
					End If
				Next
			Next

			' MsgBox "GRH - This is the list" & grh_msg_display & vbCr & "TOTAL GRH: $" & GRH_total
			PF3
		End If
	End If
	inqx_selections_has_too_many_pages = False
	If too_many_SNAP_INQX_pages = True Then inqx_selections_has_too_many_pages = True
	If too_many_MFIP_INQX_pages = True Then inqx_selections_has_too_many_pages = True
	If too_many_GA_INQX_pages = True Then inqx_selections_has_too_many_pages = True
	If too_many_MSA_INQX_pages = True Then inqx_selections_has_too_many_pages = True
	If too_many_DWP_INQX_pages = True Then inqx_selections_has_too_many_pages = True
	If too_many_GRH_INQX_pages = True Then inqx_selections_has_too_many_pages = True

	If inqx_selections_has_too_many_pages = True Then
		the_msg = "You have selected an INQX range of months that is more than 9 pages of display in MAXIS. MAXIS does not allow this display to be viewed." & vbCr & vbCr & "This is for the program:" & vbCr
		If too_many_SNAP_INQX_pages = True Then the_msg = the_msg & "  - SNAP" & vbCr
		If too_many_MFIP_INQX_pages = True Then the_msg = the_msg & "  - MFIP" & vbCr
		If too_many_GA_INQX_pages = True Then the_msg = the_msg & "  - GA" & vbCr
		If too_many_MSA_INQX_pages = True Then the_msg = the_msg & "  - MSA" & vbCr
		If too_many_DWP_INQX_pages = True Then the_msg = the_msg & "  - DWP" & vbCr
		If too_many_GRH_INQX_pages = True Then the_msg = the_msg & "  - GRH" & vbCr
		the_msg = the_msg & vbCr & "You must reduce the number of months in the range until the display of issuance is under 9 pages."
		too_many_lines_msg = MsgBox(the_msg, vbCritical, "Too Many INQX Pages")
	End If

Loop until inqx_selections_has_too_many_pages = False


'Defaulting the checkboxes for CASE specific addresses
case_address_checkbox = checked
If forms_to_arep = "Y" Then arep_address_checkbox = checked
If notc_to_swkr = "Y" Then swkr_address_checkbox = checked

'ADDRESS SELECTION Dialog
Do
	Do
		err_msg = ""
		y_pos = 25

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 551, 385, "Verification of Public Assistance"
		  ButtonGroup ButtonPressed

			Text 20, 25, 400, 10, "Check all Addresses to Include"
			GroupBox 15, 40, 450, 65, "Case Address for Mail"
			If mail_line_one <> "" Then
				Text 325, 50, 135, 10, "This case uses the MAILING address."
				Text 25, 55, 300, 10, "Adressee:" & case_name
			    Text 25, 65, 300, 10, "Street: " & mail_line_one & " " & mail_line_two
			    Text 25, 75, 70, 10, "City: " & mail_city
			    Text 145, 75, 100, 10, "State: " & mail_state
			    Text 295, 75, 70, 10, "Zip: " & mail_zip
			Else
				Text 325, 50, 135, 10, "This case uses the RESIDENCE address."
				Text 25, 55, 300, 10, "Adressee: " & case_name
			    Text 25, 65, 300, 10, "Street: " & resi_line_one & " " & resi_line_two
			    Text 25, 75, 70, 10, "City: " & resi_city
			    Text 145, 75, 100, 10, "State: " & resi_state
			    Text 295, 75, 70, 10, "Zip: " & resi_zip
			End If
		    CheckBox 20, 90, 150, 10, "Check here to mail to the Case Address", case_address_checkbox

			If arep_name <> "" Then
			    GroupBox 15, 110, 450, 65, "AREP Address"
				Text 25, 125, 300, 10, "Adressee: " & arep_name
			    Text 25, 135, 300, 10, "Street: " & arep_addr_street
			    Text 25, 145, 70, 10, "City: " & arep_addr_city
			    Text 145, 145, 100, 10, "State: " & arep_addr_state
			    Text 295, 145, 70, 10, "Zip: " & arep_addr_zip
			    CheckBox 20, 160, 150, 10, "Check here to mail to the AREP Address", arep_address_checkbox
			Else
				GroupBox 15, 110, 450, 65, "AREP Address"
				Text 25, 125, 300, 10, "No AREP Panels exists for this case."
			End If

			If swkr_name <> "" Then
			    GroupBox 15, 180, 450, 65, "Social Worker Address"
				Text 25, 195, 300, 10, "Adressee: " & swkr_name
			    Text 25, 205, 300, 10, "Street: " & swkr_addr_street
			    Text 25, 215, 70, 10, "City: " & swkr_addr_city
			    Text 145, 215, 100, 10, "State: " & swkr_addr_state
			    Text 295, 215, 70, 10, "Zip: " & swkr_addr_zip
			    CheckBox 20, 230, 150, 10, "Check here to mail to the SWKR Address", swkr_address_checkbox
			Else
				GroupBox 15, 180, 450, 65, "Social Worker Address"
				Text 25, 195, 300, 10, "No SWKR Panels exists for this case."
			End If
			GroupBox 15, 250, 450, 65, "Other Address"
			Text 25, 265, 40, 10, "Adressee:"
			EditBox 60, 260, 155, 15, other_address_person
		    Text 225, 265, 25, 10, "Street:"
			EditBox 250, 260, 205, 15, other_address_street
		    Text 25, 285, 20, 10, "City:"
			EditBox 45, 280, 95, 15, other_address_city
		    Text 170, 285, 20, 10, "State:"
			DropListBox 190, 280, 95, 45, "Select One..."+chr(9)+state_list, other_address_state
		    Text 305, 285, 15, 10, "Zip: "
			EditBox 320, 280, 75, 15, other_address_zip
		    CheckBox 20, 300, 150, 10, "Check here to mail to this Other Address", other_address_checkbox

			OkButton 445, 365, 50, 15
			CancelButton 495, 365, 50, 15
			PushButton 35, 345, 25, 10, "CURR", CURR_button
		    PushButton 60, 345, 25, 10, "PERS", PERS_button
		    PushButton 85, 345, 25, 10, "NOTE", NOTE_button
		    PushButton 160, 345, 25, 10, "XFER", XFER_button
		    PushButton 185, 345, 25, 10, "WCOM", WCOM_button
		    PushButton 210, 345, 25, 10, "MEMO", MEMO_button
		    PushButton 35, 355, 25, 10, "PROG", PROG_button
		    PushButton 60, 355, 25, 10, "MEMB", MEMB_button
		    PushButton 85, 355, 25, 10, "REVW", REVW_button
		    PushButton 160, 355, 25, 10, "INQB", INQB_button
		    PushButton 185, 355, 25, 10, "INQD", INQD_button
		    PushButton 210, 355, 25, 10, "INQX", INQX_button
		    PushButton 35, 365, 25, 10, "SNAP", ELIG_FS_button
		    PushButton 60, 365, 25, 10, "MFIP", ELIG_MFIP_button
		    PushButton 85, 365, 25, 10, "DWP", ELIG_DWP_button
		    PushButton 110, 365, 25, 10, "GA", ELIG_GA_button
		    PushButton 135, 365, 25, 10, "MSA", ELIG_MSA_button
		    PushButton 160, 365, 25, 10, "GRH", ELIG_GRH_button
		    PushButton 185, 365, 25, 10, "HC", ELIG_HC_button
		    PushButton 210, 365, 25, 10, "SUMM", ELIG_SUMM_button
		    PushButton 235, 365, 25, 10, "DENY", ELIG_DENY_button
			Text 250, 5, 290, 10, "NOTICE Information for Verification of Public Assistance for Case # " & MAXIS_case_number
			GroupBox 5, 15, 470, 315, "Details"
			GroupBox 5, 335, 390, 45, "Navigation"
			Text 10, 345, 25, 10, "CASE/"
			Text 135, 345, 25, 10, "SPEC/"
			Text 10, 355, 25, 10, "STAT/"
			Text 10, 365, 20, 10, "ELIG/"
			Text 135, 355, 25, 10, "MONY/"
	 	EndDialog

		dialog Dialog1
		cancel_confirmation
		MAXIS_dialog_navigation

		other_address_person = trim(other_address_person)		'take out the spaces
		other_address_street = trim(other_address_street)
		other_address_city = trim(other_address_city)
		other_address_state = trim(other_address_state)
		other_address_zip = trim(other_address_zip)

		'error handling for the dialog
		If other_address_checkbox = checked Then
			If other_address_person = "" Then err_msg = err_msg & vbNewLine & "* Since you have indicated the notice should go to another address as well, you must provide the name of the person this mail will go to."
			If other_address_street = "" Then err_msg = err_msg & vbNewLine & "* Since you have indicated the notice should go to another address as well, you must provide the street number and name for this address"
			If other_address_city = "" Then err_msg = err_msg & vbNewLine & "* Since you have indicated the notice should go to another address as well, you must provide the city for this address."
			If other_address_state = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since you have indicated the notice should go to another address as well, you must provide the state for this address"
			If other_address_zip =  "" Then err_msg = err_msg & vbNewLine & "* Since you have indicated the notice should go to another address as well, you must provide the zip code for this address"
		End If

	 	If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'saving information for error output email
script_run_lowdown = script_run_lowdown & vbCr & vbCr & "ADDRESS Selections:"
If case_address_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Case Address checkbox was checked"
If swkr_address_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Social Worker Address checkbox was checked"
If arep_address_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Authorized Rep Address checkbox was checked"
If other_address_checkbox = checked Then
	script_run_lowdown = script_run_lowdown & vbCr & "Other Address checkbox was checked"
	script_run_lowdown = script_run_lowdown & vbCr & "Person - " & other_address_person & vbCr & "Street - " & other_address_street & vbCr & "City - " & other_address_city & vbCr & "State - " & other_address_state & vbCr & "Zip - " & other_address_zip
End If

'setting the information for the function to send notices
If swkr_address_checkbox = unchecked Then forms_to_swkr = "N"
If arep_address_checkbox = unchecked Then forms_to_arep = "N"
If other_address_checkbox = unchecked Then send_to_other = "N"
If swkr_address_checkbox = checked Then forms_to_swkr = "Y"
If arep_address_checkbox = checked Then forms_to_arep = "Y"
If other_address_checkbox = checked Then
	send_to_other = "Y"
	other_address_state = left(other_address_state, 2)
End If

snap_resent_wcom = False		'defaults to see if the wcom is susccessful.
ga_resent_wcom = False
msa_resent_wcom = False
mfip_wcom_row = False
grh_wcom_row = False

'If any program has a WCOM to resend as the option - we are going to send it
If resend_wcom = True Then
	If snap_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(snap_month, snap_year, snap_wcom_row, snap_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
		STATS_manualtime = STATS_manualtime + 15
		' MsgBox snap_resent_wcom
	End If
	If ga_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(ga_month, ga_year, ga_wcom_row, ga_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
		STATS_manualtime = STATS_manualtime + 15
	End If
	If msa_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(msa_month, msa_year, msa_wcom_row, msa_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
		STATS_manualtime = STATS_manualtime + 15
	End If
	If mfip_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(mfip_month, mfip_year, mfip_wcom_row, mfip_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
		STATS_manualtime = STATS_manualtime + 15
	End If
	If dwp_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(dwp_month, dwp_year, dwp_wcom_row, dwp_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
	End If
	If grh_verification_method = "Resend WCOM - Eligibility Notice" Then
		Call resend_existing_wcom(grh_month, grh_year, grh_wcom_row, grh_resent_wcom, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip)
		Call back_to_SELF
	End If
End If

Call back_to_SELF		'resent

If create_memo = True Then		'If there are any MEMOs needed we need to read INQX for all the specified programs and dates and create arrays of the benefit months for each program

	'NOW we create a whole array of the lines of each possible MEMO.
	'We do it this way so we know how long each MEMO is so that we can combine Programs into a single MEMO as it best fits.
	If snap_verification_method = "Create New MEMO with range of Months" Then
		snap_array_of_memo_lines = "SNAP / Food Support Benefit:"

		for each ordered_date in SNAP_dates_array
			For each_known_issuance = 0 to UBound(SNAP_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, SNAP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					snap_array_of_memo_lines = snap_array_of_memo_lines & "~" & "   " & SNAP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		snap_array_of_memo_lines = snap_array_of_memo_lines & "~" & "SNAP Food Total for " & snap_start_month & " to " & snap_end_month & ": $" & SNAP_total
		snap_array_of_memo_lines = split(snap_array_of_memo_lines, "~")
	End If

	If ga_verification_method = "Create New MEMO with range of Months" Then
		ga_array_of_memo_lines = "GA (General Assistance) Benefit: (CASH Benefit)"
		for each ordered_date in GA_dates_array
			For each_known_issuance = 0 to UBound(GA_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, GA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					ga_array_of_memo_lines = ga_array_of_memo_lines & "~" & "   " & GA_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		ga_array_of_memo_lines = ga_array_of_memo_lines & "~" & "GA Cash Total for " & ga_start_month & " to " & ga_end_month & ": $" & GA_total
		ga_array_of_memo_lines = split(ga_array_of_memo_lines, "~")
	End If
	If msa_verification_method = "Create New MEMO with range of Months" Then
		msa_array_of_memo_lines = "MSA (MN Supplemental Aid) Benefit: (CASH Benefit)"
		for each ordered_date in MSA_dates_array
			For each_known_issuance = 0 to UBound(MSA_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, MSA_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					msa_array_of_memo_lines = msa_array_of_memo_lines & "~" & "   " & MSA_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		msa_array_of_memo_lines = msa_array_of_memo_lines & "~" & "MSA Cash Total for " & msa_start_month & " to " & msa_end_month & ": $" & MSA_total
		msa_array_of_memo_lines = split(msa_array_of_memo_lines, "~")
	End If
	If mfip_verification_method = "Create New MEMO with range of Months" Then
		mfip_array_of_memo_lines = "MFIP (MN Family Investment Program) or TANF Benefits:"
		for each ordered_date in MFIP_dates_array
			For each_known_issuance = 0 to UBound(MFIP_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, MFIP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					mfip_array_of_memo_lines = mfip_array_of_memo_lines & "~" & "   " & MFIP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		mfip_array_of_memo_lines = mfip_array_of_memo_lines & "~" & "MFIP Cash Total for " & mfip_start_month & " to " & mfip_end_month & ": $" & MFIP_Cash_total
		mfip_array_of_memo_lines = mfip_array_of_memo_lines & "~" & "MFIP Food Total for " & mfip_start_month & " to " & mfip_end_month & ": $" & MFIP_Food_total
		mfip_array_of_memo_lines = split(mfip_array_of_memo_lines, "~")
	End If
	If dwp_verification_method = "Create New MEMO with range of Months" Then
		dwp_array_of_memo_lines = "DWP Benefits:"

		for each ordered_date in DWP_dates_array
			For each_known_issuance = 0 to UBound(DWP_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, DWP_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					dwp_array_of_memo_lines = dwp_array_of_memo_lines & "~" & "   " & DWP_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		dwp_array_of_memo_lines = dwp_array_of_memo_lines & "~" & "DWP Cash Total for " & dwp_start_month & " to " & dwp_end_month & ": $" & DWP_total
		dwp_array_of_memo_lines = split(dwp_array_of_memo_lines, "~")
	End If
	If grh_verification_method = "Create New MEMO with range of Months" Then
		grh_array_of_memo_lines = "GRH / Housing Support Benefit:"

		for each ordered_date in GRH_dates_array
			For each_known_issuance = 0 to UBound(GRH_ISSUANCE_ARRAY, 2)
				If DateDiff("d", ordered_date, GRH_ISSUANCE_ARRAY(benefit_month_as_date_const, each_known_issuance)) = 0 Then
					grh_array_of_memo_lines = grh_array_of_memo_lines & "~" & "   " & GRH_ISSUANCE_ARRAY(note_message_const, each_known_issuance)
					STATS_manualtime = STATS_manualtime + 20
				End If
			Next
		Next
		grh_array_of_memo_lines = grh_array_of_memo_lines & "~" & "GRH/Housing Support Total for " & grh_start_month & " to " & grh_end_month & ": $" & GRH_total
		grh_array_of_memo_lines = split(grh_array_of_memo_lines, "~")
	End If


	'Copunting the lines in each program's MEMO
	snap_memo_lines = 0
	ga_memo_lines = 0
	msa_memo_lines = 0
	mfip_memo_lines = 0
	dwp_memo_lines = 0
	memo_count = 0
	Dim EACH_MEMO_ARRAY()
	ReDim EACH_MEMO_ARRAY(0)
	If IsArray(snap_array_of_memo_lines) = True Then
		snap_memo_lines = UBound(snap_array_of_memo_lines) + 1
	End If
	If IsArray(ga_array_of_memo_lines) = True Then
		ga_memo_lines = UBound(ga_array_of_memo_lines) + 1
	End If
	If IsArray(msa_array_of_memo_lines) = True Then
		msa_memo_lines = UBound(msa_array_of_memo_lines) + 1
	End If
	If IsArray(mfip_array_of_memo_lines) = True Then
		mfip_memo_lines = UBound(mfip_array_of_memo_lines) + 1
	End If
	If IsArray(dwp_array_of_memo_lines) = True Then
		dwp_memo_lines = UBound(dwp_array_of_memo_lines) + 1
	End If
	If IsArray(grh_array_of_memo_lines) = True Then
		grh_memo_lines = UBound(grh_array_of_memo_lines) + 1
	End If

	'Now we have a lot of logic to try to combine the counts to fit into a MEMO to attempt to send as few MEMOs as possible.
	'MOST of the time we will only have 1 MEMO
	total_memo_lines = snap_memo_lines + ga_memo_lines + msa_memo_lines + mfip_memo_lines + dwp_memo_lines + grh_memo_lines + 3		'Best option - all programs on one MEMO
	memo_list = ""
	need_cover_memo = False
	If total_memo_lines < 28 Then
		memo_list = memo_list & "~MFIP/GA/MSA/SNAP/DWP/GRH"
	Else
		need_cover_memo	= True																		'If they aren't all on one, then we add a memo
		If mfip_memo_lines > 0 Then																'These are for if any program needs more than one on their own
			memo_list = memo_list & "~MFIP"
			If mfip_memo_lines > 27 Then memo_list = memo_list & "~MFIP"																'These are for if any program needs more than one on their own
			If mfip_memo_lines > 55 Then memo_list = memo_list & "~MFIP"
			If mfip_memo_lines > 83 Then memo_list = memo_list & "~MFIP"
		End If
		If snap_memo_lines > 0 Then
			memo_list = memo_list & "~SNAP"
			If snap_memo_lines > 27 Then memo_list = memo_list & "~SNAP"
			If snap_memo_lines > 55 Then memo_list = memo_list & "~SNAP"
			If snap_memo_lines > 83 Then memo_list = memo_list & "~SNAP"
		End If
		If dwp_memo_lines > 0 Then
			memo_list = memo_list & "~DWP"
			If dwp_memo_lines > 27 Then memo_list = memo_list & "~DWP"
			If dwp_memo_lines > 55 Then memo_list = memo_list & "~DWP"
			If dwp_memo_lines > 83 Then memo_list = memo_list & "~DWP"
		End If
		If ga_memo_lines > 0 Then
			memo_list = memo_list & "~GA"
			If ga_memo_lines > 27 Then memo_list = memo_list & "~GA"
			If ga_memo_lines > 55 Then memo_list = memo_list & "~GA"
			If ga_memo_lines > 83 Then memo_list = memo_list & "~GA"
		End If
		If msa_memo_lines > 0 Then
			memo_list = memo_list & "~MSA"
			If msa_memo_lines > 27 Then memo_list = memo_list & "~MSA"
			If msa_memo_lines > 55 Then memo_list = memo_list & "~MSA"
			If msa_memo_lines > 83 Then memo_list = memo_list & "~MSA"
		End If
		If grh_memo_lines > 0 Then
			memo_list = memo_list & "~GRH"
			If grh_memo_lines > 27 Then memo_list = memo_list & "~GRH"
			If grh_memo_lines > 55 Then memo_list = memo_list & "~GRH"
			If grh_memo_lines > 83 Then memo_list = memo_list & "~GRH"
		End If

		If mfip_memo_lines + ga_memo_lines + msa_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine four programs into one MEMO
			memo_list = memo_list & "~MFIP/GA/MSA/GRH/DWP"
			If snap_memo_lines <> 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
		ElseIf snap_memo_lines + ga_memo_lines + msa_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MSA/GRH/DWP"
			If mfip_memo_lines <> 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
		ElseIf snap_memo_lines + ga_memo_lines + mfip_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MFIP/GRH/DWP"
			If msa_memo_lines <> 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
		ElseIf snap_memo_lines + mfip_memo_lines + msa_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/MSA/MFIP/GRH/DWP"
			If ga_memo_lines <> 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
		ElseIf mfip_memo_lines + ga_memo_lines + msa_memo_lines + snap_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~SNAP/MFIP/GA/MSA/DWP"
			If grh_memo_lines <> 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
		ElseIf mfip_memo_lines + ga_memo_lines + msa_memo_lines + snap_memo_lines + grh_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~SNAP/MFIP/GA/MSA/GRH"
			If dwp_memo_lines <> 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"

		ElseIf mfip_memo_lines + ga_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~MFIP/GA/MSA/DWP"
			If grh_memo_lines + snap_memo_lines < 28 Then
				memo_list = memo_list & "~SNAP/GRH"
			Else
				If snap_memo_lines > 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
				If grh_memo_lines > 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
			End If
		ElseIf mfip_memo_lines + ga_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~MFIP/GA/MSA/GRH"
			If dwp_memo_lines + snap_memo_lines < 28 Then
				memo_list = memo_list & "~SNAP/DWP"
			Else
				If snap_memo_lines > 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
				If dwp_memo_lines > 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
			End If
		ElseIf mfip_memo_lines + ga_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~MFIP/GA/GRH/DWP"
			If msa_memo_lines + snap_memo_lines < 28 Then
				memo_list = memo_list & "~SNAP/MSA"
			Else
				If snap_memo_lines > 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
				If msa_memo_lines > 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
			End If
		ElseIf mfip_memo_lines + grh_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~MFIP/GRH/MSA/DWP"
			If ga_memo_lines + snap_memo_lines < 28 Then
				memo_list = memo_list & "~SNAP/GA"
			Else
				If snap_memo_lines > 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
				If ga_memo_lines > 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
			End If
		ElseIf grh_memo_lines + ga_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then							'These try to combine three programs into one MEMO
			memo_list = memo_list & "~GRH/GA/MSA/DWP"
			If mfip_memo_lines + snap_memo_lines < 28 Then
				memo_list = memo_list & "~SNAP/MFIP"
			Else
				If snap_memo_lines > 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
				If mfip_memo_lines > 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
			End If


		ElseIf snap_memo_lines + ga_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MSA/DWP"
			If grh_memo_lines + mfip_memo_lines < 28 Then
				memo_list = memo_list & "~MFIP/GRH"
			Else
				If mfip_memo_lines > 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
				If grh_memo_lines > 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
			End If
		ElseIf snap_memo_lines + grh_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GRH/MSA/DWP"
			If ga_memo_lines + mfip_memo_lines < 28 Then
				memo_list = memo_list & "~MFIP/GA"
			Else
				If mfip_memo_lines > 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
				If ga_memo_lines > 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
			End If
		ElseIf snap_memo_lines + ga_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/GRH/DWP"
			If msa_memo_lines + mfip_memo_lines < 28 Then
				memo_list = memo_list & "~MFIP/MSA"
			Else
				If mfip_memo_lines > 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
				If msa_memo_lines > 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
			End If
		ElseIf snap_memo_lines + ga_memo_lines + msa_memo_lines + grh_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MSA/GRH"
			If dwp_memo_lines + mfip_memo_lines < 28 Then
				memo_list = memo_list & "~MFIP/DWP"
			Else
				If mfip_memo_lines > 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"
				If dwp_memo_lines > 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
			End If

		ElseIf snap_memo_lines + mfip_memo_lines + msa_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/MSA/MFIP/DWP"
			If grh_memo_lines + ga_memo_lines < 28 Then
				memo_list = memo_list & "~GA/GRH"
			Else
				If ga_memo_lines > 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
				If grh_memo_lines > 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
			End If
		ElseIf snap_memo_lines + mfip_memo_lines + grh_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GRH/MFIP/DWP"
			If msa_memo_lines + ga_memo_lines < 28 Then
				memo_list = memo_list & "~GA/GRH"
			Else
				If ga_memo_lines > 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
				If msa_memo_lines > 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
			End If
		ElseIf snap_memo_lines + mfip_memo_lines + msa_memo_lines + grh_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/MSA/MFIP/GRH"
			If dwp_memo_lines + ga_memo_lines < 28 Then
				memo_list = memo_list & "~GA/DWP"
			Else
				If ga_memo_lines > 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
				If dwp_memo_lines > 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
			End If

		ElseIf snap_memo_lines + ga_memo_lines + mfip_memo_lines + dwp_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MFIP/DWP"
			If grh_memo_lines + msa_memo_lines < 28 Then
				memo_list = memo_list & "~MSA/GRH"
			Else
				If msa_memo_lines > 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
				If grh_memo_lines > 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
			End If
		ElseIf snap_memo_lines + ga_memo_lines + mfip_memo_lines + grh_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/GA/MFIP/GRH"
			If dwp_memo_lines + msa_memo_lines < 28 Then
				memo_list = memo_list & "~MSA/DWP"
			Else
				If msa_memo_lines > 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
				If dwp_memo_lines > 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
			End If

		ElseIf snap_memo_lines + mfip_memo_lines + msa_memo_lines + ga_memo_lines + 2 < 28 Then
			memo_list = memo_list & "~SNAP/MSA/MFIP/GA"
			If grh_memo_lines + dwp_memo_lines < 28 Then
				memo_list = memo_list & "~DWP/GRH"
			Else
				If dwp_memo_lines > 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
				If grh_memo_lines > 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
			End If
		
		Else
			If mfip_memo_lines <> 0 AND mfip_memo_lines < 28 Then memo_list = memo_list & "~MFIP"		'These are just single program MEMOs
			If ga_memo_lines <> 0 AND ga_memo_lines < 28 Then memo_list = memo_list & "~GA"
			If msa_memo_lines <> 0 AND msa_memo_lines < 28 Then memo_list = memo_list & "~MSA"
			If snap_memo_lines <> 0 AND snap_memo_lines < 28 Then memo_list = memo_list & "~SNAP"
			If dwp_memo_lines <> 0 AND dwp_memo_lines < 28 Then memo_list = memo_list & "~DWP"
			If grh_memo_lines <> 0 AND grh_memo_lines < 28 Then memo_list = memo_list & "~GRH"
		End If

	End If

	If left(memo_list, 1) = "~" Then memo_list = right(memo_list, len(memo_list) - 1)		'Making the list of the MEMOs by program an actual ARRAY'

	If InStr(memo_list, "~") = 0 Then
		the_memos_array = Array(memo_list)
	Else
		the_memos_array = split(memo_list, "~")
	End If

	snap_restart_memo_lines_position = 0										'Starting values for where we begin to count
	ga_restart_memo_lines_position = 0
	msa_restart_memo_lines_position = 0
	mfip_restart_memo_lines_position = 0
	grh_restart_memo_lines_position = 0
	dwp_restart_memo_lines_position = 0

	For each memo_to_write in the_memos_array									'Using the Array we just made, we care going to make 1 MEMO for each identified program
		memo_line = 1
		memo_full = False
		'New function with all the options for starting a MEMO
		Call start_a_new_spec_memo(memo_opened, False, forms_to_arep, forms_to_swkr, send_to_other, other_address_person, other_address_street, other_address_city, other_address_state, other_address_zip, False)

		If memo_full = False AND snap_verification_method = "Create New MEMO with range of Months" AND InStr(memo_to_write, "SNAP") <> 0 Then
			If snap_restart_memo_lines_position <= UBound(snap_array_of_memo_lines) Then
				If memo_line = 1 AND snap_restart_memo_lines_position = 0 Then											'at the becinnging of the program and the MEMO
					Call write_variable_in_SPEC_MEMO("               Benefit Issuances by Month")
					Call write_variable_in_SPEC_MEMO("                           Information provided by request.")
					memo_line = memo_line + 2
				ElseIf memo_line <> 1 AND snap_restart_memo_lines_position = 0 Then										'in the middle of a MEMO, but the beginning of a program'
					Call write_variable_in_SPEC_MEMO("-      - - - - - - - - - - - - - - - - - - - -       -")
					memo_line = memo_line + 1
				ElseIf snap_restart_memo_lines_position <> 0 Then														'In the middle of a program and the beginning of a MEMO'
					Call write_variable_in_SPEC_MEMO("SNAP (Food Support) Benefit Information CONTINUED:")
					memo_line = memo_line + 1
				End If
				For snap_info = snap_restart_memo_lines_position to UBound(snap_array_of_memo_lines)					'Now write the lines we created in the INQX functionality
					Call write_variable_in_SPEC_MEMO(snap_array_of_memo_lines(snap_info))
					end_line = snap_info + 1
					memo_line = memo_line + 1
					If memo_line = 30 Then		'We stop at 30 because there is a line to add to line 30 for every MEMO'
						memo_full = True
						Exit For
					End If
				Next
				snap_restart_memo_lines_position = end_line			'Setting where we ended for this program'
			End If
		End If
		If memo_full = False AND ga_verification_method = "Create New MEMO with range of Months" AND InStr(memo_to_write, "GA") <> 0 Then
			If ga_restart_memo_lines_position <= UBound(ga_array_of_memo_lines) Then
				If memo_line = 1 AND ga_restart_memo_lines_position = 0 Then											'at the becinnging of the program and the MEMO
					Call write_variable_in_SPEC_MEMO("               Benefit Issuances by Month")
					Call write_variable_in_SPEC_MEMO("                           Information provided by request.")
					memo_line = memo_line + 2
				ElseIf memo_line <> 1 AND ga_restart_memo_lines_position = 0 Then										'in the middle of a MEMO, but the beginning of a program'
					Call write_variable_in_SPEC_MEMO("-      - - - - - - - - - - - - - - - - - - - -       -")
					memo_line = memo_line + 1
				ElseIf ga_restart_memo_lines_position <> 0 Then														'In the middle of a program and the beginning of a MEMO'
					Call write_variable_in_SPEC_MEMO("GA Benefit Information CONTINUED:")
					memo_line = memo_line + 1
				End If
				For ga_info = ga_restart_memo_lines_position to UBound(ga_array_of_memo_lines)					'Now write the lines we created in the INQX functionality
					Call write_variable_in_SPEC_MEMO(ga_array_of_memo_lines(ga_info))
					end_line = ga_info + 1
					memo_line = memo_line + 1
					If memo_line = 30 Then		'We stop at 30 because there is a line to add to line 30 for every MEMO'
						memo_full = True
						Exit For
					End If
				Next
				ga_restart_memo_lines_position = end_line			'Setting where we ended for this program
			End If
		End If
		If memo_full = False AND msa_verification_method = "Create New MEMO with range of Months" AND InStr(memo_to_write, "MSA") <> 0 then
			If msa_restart_memo_lines_position <= UBound(msa_array_of_memo_lines) Then
				If memo_line = 1 AND msa_restart_memo_lines_position = 0 Then											'at the becinnging of the program and the MEMO
					Call write_variable_in_SPEC_MEMO("               Benefit Issuances by Month")
					Call write_variable_in_SPEC_MEMO("                           Information provided by request.")
					memo_line = memo_line + 2
				ElseIf memo_line <> 1 AND msa_restart_memo_lines_position = 0 Then										'in the middle of a MEMO, but the beginning of a program'
					Call write_variable_in_SPEC_MEMO("-      - - - - - - - - - - - - - - - - - - - -       -")
					memo_line = memo_line + 1
				ElseIf msa_restart_memo_lines_position <> 0 Then														'In the middle of a program and the beginning of a MEMO'
					Call write_variable_in_SPEC_MEMO("MSA Benefit Information CONTINUED:")
					memo_line = memo_line + 1
				End If
				For msa_info = msa_restart_memo_lines_position to UBound(msa_array_of_memo_lines)					'Now write the lines we created in the INQX functionality
					Call write_variable_in_SPEC_MEMO(msa_array_of_memo_lines(msa_info))
					end_line = msa_info + 1
					memo_line = memo_line + 1
					If memo_line = 30 Then		'We stop at 30 because there is a line to add to line 30 for every MEMO'
						memo_full = True
						Exit For
					End If
				Next
				msa_restart_memo_lines_position = end_line			'Setting where we ended for this program
			End If
		End If
		If memo_full = False AND mfip_verification_method = "Create New MEMO with range of Months" AND InStr(memo_to_write, "MFIP") <> 0 Then
			If mfip_restart_memo_lines_position <= UBound(mfip_array_of_memo_lines) Then
				If memo_line = 1 AND mfip_restart_memo_lines_position = 0 Then											'at the becinnging of the program and the MEMO
					Call write_variable_in_SPEC_MEMO("               Benefit Issuances by Month")
					Call write_variable_in_SPEC_MEMO("                           Information provided by request.")
					memo_line = memo_line + 2
				ElseIf memo_line <> 1 AND mfip_restart_memo_lines_position = 0 Then										'in the middle of a MEMO, but the beginning of a program'
					Call write_variable_in_SPEC_MEMO("-      - - - - - - - - - - - - - - - - - - - -       -")
					memo_line = memo_line + 1
				ElseIf mfip_restart_memo_lines_position <> 0 Then														'In the middle of a program and the beginning of a MEMO'
					Call write_variable_in_SPEC_MEMO("MFIP Benefit Information CONTINUED:")
					memo_line = memo_line + 1
				End If
				For mfip_info = mfip_restart_memo_lines_position to UBound(mfip_array_of_memo_lines)					'Now write the lines we created in the INQX functionality
					Call write_variable_in_SPEC_MEMO(mfip_array_of_memo_lines(mfip_info))
					end_line = mfip_info + 1
					memo_line = memo_line + 1
					If memo_line = 30 Then		'We stop at 30 because there is a line to add to line 30 for every MEMO'
						memo_full = True
						Exit For
					End If
				Next
				mfip_restart_memo_lines_position = end_line			'Setting where we ended for this program
			End If
		End If
		If memo_full = False AND grh_verification_method = "Create New MEMO with range of Months" AND InStr(memo_to_write, "GRH") <> 0 Then
			If grh_restart_memo_lines_position <= UBound(grh_array_of_memo_lines) Then
				If memo_line = 1 AND grh_restart_memo_lines_position = 0 Then											'at the becinnging of the program and the MEMO
					Call write_variable_in_SPEC_MEMO("               Benefit Issuances by Month")
					Call write_variable_in_SPEC_MEMO("                           Information provided by request.")
					memo_line = memo_line + 2
				ElseIf memo_line <> 1 AND grh_restart_memo_lines_position = 0 Then										'in the middle of a MEMO, but the beginning of a program'
					Call write_variable_in_SPEC_MEMO("-      - - - - - - - - - - - - - - - - - - - -       -")
					memo_line = memo_line + 1
				ElseIf grh_restart_memo_lines_position <> 0 Then														'In the middle of a program and the beginning of a MEMO'
					Call write_variable_in_SPEC_MEMO("GRH (Housing Support) Benefit Information CONTINUED:")
					memo_line = memo_line + 1
				End If
				For grh_info = grh_restart_memo_lines_position to UBound(grh_array_of_memo_lines)					'Now write the lines we created in the INQX functionality
					Call write_variable_in_SPEC_MEMO(grh_array_of_memo_lines(grh_info))
					end_line = grh_info + 1
					memo_line = memo_line + 1
					If memo_line = 30 Then		'We stop at 30 because there is a line to add to line 30 for every MEMO'
						memo_full = True
						Exit For
					End If
				Next
				grh_restart_memo_lines_position = end_line			'Setting where we ended for this program'
			End If
		End If
		Do while memo_line < 30										'Now we want to get to line 30
			Call write_variable_in_SPEC_MEMO("")
			memo_line = memo_line + 1
		Loop
		Call write_variable_in_SPEC_MEMO("This information is accurate and complete as of " & date)		'This goes on line 30

		PF4

		'SAVE THIS for TESTING - we can 'uncomment' and comment out the PF4 so that MEMOs are not created - helpful for testing and training
		' MsgBox "MEMO Done " & vbCr & memo_to_write
		' PF3
		' PF3
		' MsgBox "confirm erased"
	Next
	Call back_to_SELF		'reset'
End If

pa_verif_programs = ""		'Lets make a string of all the programs addressed

If snap_resent_wcom = True OR snap_verification_method = "Create New MEMO with range of Months" OR snap_not_actv_memo_for_old_beneftis_checkbox = checked Then pa_verif_programs = pa_verif_programs & "/SNAP"
If ga_resent_wcom = True OR ga_verification_method = "Create New MEMO with range of Months" OR ga_not_actv_memo_for_old_beneftis_checkbox = checked Then pa_verif_programs = pa_verif_programs & "/GA"
If msa_resent_wcom = True OR msa_verification_method = "Create New MEMO with range of Months" OR msa_not_actv_memo_for_old_beneftis_checkbox = checked Then pa_verif_programs = pa_verif_programs & "/MSA"
If mfip_resent_wcom = True OR mfip_verification_method = "Create New MEMO with range of Months" OR mfip_not_actv_memo_for_old_beneftis_checkbox = checked Then pa_verif_programs = pa_verif_programs & "/MFIP"
If grh_resent_wcom = True OR grh_verification_method = "Create New MEMO with range of Months" OR grh_not_actv_memo_for_old_beneftis_checkbox = checked Then pa_verif_programs = pa_verif_programs & "/GRH"

If left(pa_verif_programs, 1) = "/" Then pa_verif_programs = right(pa_verif_programs, len(pa_verif_programs)-1)

Call start_a_blank_CASE_NOTE			'Now we are CASE:NOTING

' Call write_variable_in_CASE_NOTE("Verification of " & pa_verif_programs & " Assistance sent")
Call write_variable_in_CASE_NOTE("Verification of Public Assistance Requested")
Call write_variable_in_CASE_NOTE("Requested by: " & verif_request_by)
If snap_resent_wcom = True Then
	Call write_variable_in_CASE_NOTE("SNAP WCOM resent to Client from " & snap_month & "/" & snap_year & ".")
	Call write_variable_in_CASE_NOTE("   - " & snap_wcom_text)
End If
If snap_verification_method = "Create New MEMO with range of Months" Then Call write_variable_in_CASE_NOTE("SPEC/MEMO sent with SNAP benefits summary from " & snap_start_month & " to " &  snap_end_month & ", per INQX.")

If ga_resent_wcom = True Then
	Call write_variable_in_CASE_NOTE("GA WCOM resent to Client from " & ga_month & "/" & ga_year & ".")
	Call write_variable_in_CASE_NOTE("   - " & ga_wcom_text)
End If
If ga_verification_method = "Create New MEMO with range of Months" Then Call write_variable_in_CASE_NOTE("SPEC/MEMO sent with GA benefits summary from " & ga_start_month & " to " &  ga_end_month & ", per INQX.")

If msa_resent_wcom = True Then
	Call write_variable_in_CASE_NOTE("MSA WCOM resent to Client from " & msa_month & "/" & msa_year & ".")
	Call write_variable_in_CASE_NOTE("   - " & msa_wcom_text)
End If
If msa_verification_method = "Create New MEMO with range of Months" Then Call write_variable_in_CASE_NOTE("SPEC/MEMO sent with MSA benefits summary from " & msa_start_month & " to " &  msa_end_month & ", per INQX.")

If mfip_resent_wcom = True Then
	Call write_variable_in_CASE_NOTE("MFIP WCOM resent to Client from " & mfip_month & "/" & mfip_year & ".")
	Call write_variable_in_CASE_NOTE("   - " & mfip_wcom_text)
End If
If mfip_verification_method = "Create New MEMO with range of Months" Then Call write_variable_in_CASE_NOTE("SPEC/MEMO sent with MFIP benefits summary from " & mfip_start_month & " to " &  mfip_end_month & ", per INQX.")

If grh_resent_wcom = True Then
	Call write_variable_in_CASE_NOTE("GRH WCOM resent to Client from " & grh_month & "/" & grh_year & ".")
	Call write_variable_in_CASE_NOTE("   - " & grh_wcom_text)
End If
If grh_verification_method = "Create New MEMO with range of Months" Then Call write_variable_in_CASE_NOTE("SPEC/MEMO sent with GRH benefits summary from " & grh_start_month & " to " &  grh_end_month & ", per INQX.")

If forms_to_arep = "Y" Then Call write_variable_in_CASE_NOTE("Notices sent to AREP.")
If forms_to_swkr = "Y" Then Call write_variable_in_CASE_NOTE("Notices sent to SWKR.")
If send_to_other = "Y" Then
	Call write_variable_in_CASE_NOTE("Notices sent to address provided by client.")
	Call write_variable_in_CASE_NOTE("   " & other_address_street)
	Call write_variable_in_CASE_NOTE("   " & other_address_city & ", " & other_address_state & " " & other_address_zip)
End If
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Notice sent for PA Verif Request")
'THIS IS WHERE I STOPPED ADDING DWP _ START HERE'
