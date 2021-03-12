'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - RETURNED MAIL RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if'Required for statistical purposes==========================================================================================
		name_of_script = "NOTES - RETURNED MAIL RECEIVED.vbs"
		start_time = timer
		STATS_counter = 1               'sets the stats counter at one
		STATS_manualtime = 360          'manual run time in seconds
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

		'CHANGELOG BLOCK ===========================================================================================================
		'Starts by defining a changelog array
		changelog = array()

		'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
		'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
		call changelog_update("03/12/2021", "GitHub Issue #309 Updated handling for current address confirmation.", "MiKayla Handley")
		call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
		call changelog_update("02/13/2020", "Updated the zip code to only allow for 5 characters.", "MiKayla Handley, Hennepin County")
		call changelog_update("06/06/2019", "Initial version. Re-written per POLI/TEMP.", "MiKayla Handley, Hennepin County")

		'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
		changelog_display

		'END CHANGELOG BLOCK =======================================================================================================
		FUNCTION read_ADDR_panel(addr_eff_date, addr_future_date, resi_addr_line_one, resi_addr_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, mail_line_one, mail_line_two, mail_city_line, mail_state_line, mail_zip_line, living_situation, living_sit_line, homeless_line, addr_phone_1A)

		    Call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen addr_eff_date, 8, 4, 43 'todo build in more handling for future dates '
			EMReadScreen addr_future_date, 8, 4, 66
			EMreadscreen resi_addr_line_one, 22, 6, 43
			EMreadscreen resi_addr_line_two, 22, 7, 43
			EMreadscreen resi_addr_city, 15, 8, 43
			EMreadscreen resi_addr_state, 2, 8, 66
			EMreadscreen resi_addr_zip, 7, 9, 43
			EMreadscreen resi_addr_county, 2, 9, 66
			EMreadscreen addr_verif, 2, 9, 74
			EMreadscreen addr_homeless, 1, 10, 43
			EMreadscreen addr_reservation, 1, 10, 74
			EMreadscreen addr_phone_1A, 3, 17, 45	'Has to split phone numbers up into three parts each sometimes they are partially complete
			EMreadscreen addr_phone_2B, 3, 17, 51
			EMreadscreen addr_phone_3C, 4, 17, 55
			EMreadscreen addr_phone_2A, 3, 18, 45
			EMreadscreen addr_phone_2B, 3, 18, 51
			EMreadscreen addr_phone_2C, 4, 18, 55
			EMreadscreen addr_phone_3A, 3, 19, 45
			EMreadscreen addr_phone_3B, 3, 19, 51
			EMreadscreen addr_phone_3C, 4, 19, 55
			EMReadScreen verif_line, 2, 9, 74
			EMReadScreen homeless_line, 1, 10, 43
			EMReadScreen reservation_line, 1, 10, 74
			EMReadScreen living_sit_line, 2, 11, 43

			EMReadScreen mail_line_one, 22, 13, 43
			EMReadScreen mail_line_two, 22, 14, 43
			EMReadScreen mail_city_line, 15, 15, 43
			EMReadScreen mail_state_line, 2, 16, 43
			EMReadScreen mail_zip_line, 7, 16, 52

			addr_eff_date = replace(addr_eff_date, " ", "/")
			addr_future_date = trim(addr_future_date)
			addr_future_date = replace(addr_future_date, " ", "/")

			resi_addr_line_one = replace(resi_addr_line_one, "_", "")
		    resi_addr_line_two = replace(resi_addr_line_two, "_", "")
		    resi_addr_city = replace(resi_addr_city, "_", "")
		    resi_addr_state = replace(resi_addr_state, "_", "")
		    resi_addr_zip = replace(resi_addr_zip, "_", "")

			mail_line_one = replace(mail_line_one, "_", "")
			mail_line_two = replace(mail_line_two, "_", "")
			mail_city_line = replace(mail_city_line, "_", "")
			mail_state_line = replace(mail_state_line, "_", "")
			mail_zip_line = replace(mail_zip_line, "_", "")

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

		    If homeless_line = "Y" Then homeless_yn = "Yes"
		    If homeless_line = "N" Then homeless_yn = "No"
		    If reservation_line = "Y" Then reservation_yn = "Yes"
		    If reservation_line = "N" Then reservation_yn = "No"

		    If verif_line = "SF" Then addr_verif = "SF - Shelter Form"
		    If verif_line = "CO" Then addr_verif = "CO - Coltrl Stmt"
		    If verif_line = "MO" Then addr_verif = "MO - Mortgage Papers"
		    If verif_line = "TX" Then addr_verif = "TX - Prop Tax Stmt"
		    If verif_line = "CD" Then addr_verif = "CD - Contract for Deed"
		    If verif_line = "UT" Then addr_verif = "UT - Utility Stmt"
		    If verif_line = "DL" Then addr_verif = "DL - Driver Lic/State ID"
		    If verif_line = "OT" Then addr_verif = "OT - Other Document"
		    If verif_line = "NO" Then addr_verif = "NO - No Ver Prvd"
		    If verif_line = "?_" Then addr_verif = "? - Delayed"
		    If verif_line = "__" Then addr_verif = "Blank"

		    If living_sit_line = "__" Then living_situation = "Blank"
		    If living_sit_line = "01" Then living_situation = "01 - Own home, lease or roommate"
		    If living_sit_line = "02" Then living_situation = "02 - Family/Friends - economic hardship"
		    If living_sit_line = "03" Then living_situation = "03 - Servc prvdr- foster/group home"
		    If living_sit_line = "04" Then living_situation = "04 - Hospital/Treatment/Detox/Nursing Home"
		    If living_sit_line = "05" Then living_situation = "05 - Jail/Prison//Juvenile Det."
		    If living_sit_line = "06" Then living_situation = "06 - Hotel/Motel"
		    If living_sit_line = "07" Then living_situation = "07 - Emergency Shelter"
		    If living_sit_line = "08" Then living_situation = "08 - Place not meant for Housing"
		    If living_sit_line = "09" Then living_situation = "09 - Declined"
		    If living_sit_line = "10" Then living_situation = "10 - Unknown"

		    notes_on_address = "Address effective: " & addr_eff_date & "."
		    If mail_line_one <> "" Then
		        If mail_line_two = "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
		        If mail_line_two <> "" Then notes_on_address = notes_on_address & " Mailing address: " & mail_line_one & " " & mail_line_two & " " & mail_city_line & ", " & mail_state_line & " " & mail_zip_line
		    End If
		    If addr_future_date <> "" Then notes_on_address = notes_on_address & "; ** Address will update effective " & addr_future_date & "."
		END FUNCTION

		'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
		'Connects to BLUEZONE
		EMConnect ""
		'Grabs the MAXIS case number
		CALL MAXIS_case_number_finder(MAXIS_case_number)

		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		'Running the initial dialog
		BeginDialog Dialog1, 0, 0, 211, 85, "Returned Mail"
		  EditBox 55, 5, 40, 15, MAXIS_case_number
		  EditBox 155, 5, 50, 15, date_received
		  DropListBox 90, 25, 115, 15, "Select One:"+chr(9)+"forwarding address in MN"+chr(9)+"forwarding address outside MN"+chr(9)+"no forwarding address provided"+chr(9)+"no response received", ADDR_actions
		  EditBox 90, 45, 115, 15, worker_signature
		  ButtonGroup ButtonPressed
		    OkButton 110, 65, 45, 15
		    CancelButton 160, 65, 45, 15
		  Text 5, 30, 85, 10, "Mail Has Been Returned:"
		  Text 5, 50, 60, 10, "Worker Signature:"
		  Text 5, 10, 50, 10, "Case Number:"
		  Text 100, 10, 50, 10, "Date Received:"
		EndDialog

		DO
		    DO
		    	err_msg = ""
		    	DIALOG Dialog1
		    	cancel_confirmation
		    	IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
				IF isdate(date_received) = FALSE THEN err_msg = err_msg & vbnewline & "Please enter the date."
				IF Cdate(date_received) > cdate(date) = TRUE THEN err_msg = err_msg & vbnewline & "You must enter an actual date that is not in the future and is in the footer month that you are working in."
				IF ADDR_actions = "Select One:" THEN err_msg = err_msg & vbCr & "Please chose an action for the returned mail."
		    	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
		    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		    LOOP UNTIL err_msg = ""
		    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		'setting the footer month to make the updates in'
		CALL convert_date_into_MAXIS_footer_month(date_received, MAXIS_footer_month, MAXIS_footer_year)
		MAXIS_footer_month_confirmation

		CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
		IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

		CALL read_ADDR_panel(addr_eff_date, addr_future_date, resi_addr_line_one, resi_addr_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, mail_line_one, mail_line_two, mail_city_line, mail_state_line, mail_zip_line, living_situation, living_sit_line, homeless_line, addr_phone_1A)

		EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to navigate'
		IF trim(case_invalid_error) <> "" THEN script_end_procedure("*** NOTICE!!! ***" & vbCr & case_invalid_error & vbCr & "Please resolve for the script to continue.")
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 566, 90, "Returned Mail Information"
		  DropListBox 220, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", mailing_address_confirmed
		  DropListBox 500, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", residential_address_confirmed
		  EditBox 75, 70, 205, 15, notes_on_address
		  ButtonGroup ButtonPressed
		    PushButton 285, 70, 50, 15, "CASE/ADHI", ADHI_button
		    PushButton 340, 70, 70, 15, "HSR Manual/SNAP", HSR_manual_button
		    OkButton 450, 70, 55, 15
		    CancelButton 505, 70, 55, 15
		  Text 10, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
		  Text 15, 30, 200, 10, mail_line_one
		  Text 15, 40, 200, 10, mail_line_two
		  Text 15, 50, 205, 10, mail_city_line & ", " & mail_state_line &  ", " & mail_zip_line
		  Text 10, 75, 65, 10, "Notes on Address:"
		  GroupBox 285, 5, 275, 60, "Residential Address"
		  Text 290, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
		  Text 295, 30, 200, 10, resi_addr_line_one
		  Text 295, 40, 200, 10, resi_addr_line_two
		  Text 295, 50, 205, 10, resi_addr_city &  ", " & resi_addr_state &  ", "  & resi_addr_zip
		  GroupBox 5, 5, 275, 60, "Mailing Address in Maxis"
		EndDialog
		BeginDialog Dialog1, 0, 0, 566, 105, "Returned Mail Information"
		  DropListBox 220, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", mailing_address_confirmed
		  DropListBox 500, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", residential_address_confirmed
		  EditBox 75, 85, 205, 15, notes_on_address
		  ButtonGroup ButtonPressed
		    PushButton 285, 85, 50, 15, "CASE/ADHI", ADHI_button
		    PushButton 340, 85, 70, 15, "HSR Manual/SNAP", HSR_manual_button
		    OkButton 445, 85, 55, 15
		    CancelButton 505, 85, 55, 15
		  Text 10, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
		  Text 15, 30, 200, 10, mail_line_one
		  Text 15, 40, 200, 10, mail_line_two
		  Text 15, 50, 205, 10, mail_city_line & ", " & mail_state_line &  ", " & mail_zip_line
		  Text 10, 90, 65, 10, "Notes on Address:"
		  GroupBox 285, 5, 275, 75, "Residential Address"
		  Text 290, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
		  Text 295, 30, 200, 10, resi_addr_line_one
		  Text 295, 40, 200, 10, resi_addr_line_two
		  Text 295, 50, 205, 10, resi_addr_city &  ", " & resi_addr_state &  ", "  & resi_addr_zip
		  GroupBox 5, 5, 275, 75, "Mailing Address in Maxis"
		  Text 295, 65, 50, 10, "Effective Date:"
		  Text 15, 65, 50, 10, "Future Date:"
		  Text 350, 65, 70, 10, addr_eff_date
		  Text 70, 65, 70, 10, addr_future_date
		EndDialog

		DO
		    DO
				err_msg = ""
		    	DO
					Dialog Dialog1
					cancel_without_confirmation
					MAXIS_dialog_navigation
		            IF buttonpressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/return_Mail_Processing_for_SNAP.aspx") 'HSR manual policy page
		        LOOP until ButtonPressed = -1
				IF mailing_address_confirmed = "Select One:" THEN err_msg = err_msg & vbCr & "Please confirm if the mailing address is the address that the agency ttempted to deliver mail to."
				IF residential_address_confirmed = "Select One:" THEN err_msg = err_msg & vbCr & "Please confirm if this the address that the agency attempted to deliver mail to."
				IF mailing_address_confirmed = "NO" and residential_address_confirmed = "NO" and notes_on_address = "" THEN  err_msg = err_msg & vbCr & "Please confirm what the address was using notes on address that the agency attempted to deliver mail to. ADDR will not be updated at this time. Please explain where the address was found and on what date."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				LOOP UNTIL err_msg = ""
		    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to  assword back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		IF mailing_address_confirmed = "YES" or residential_address_confirmed = "YES" THEN
		    '-------------------------------------------------------------------------------------------------DIALOG
		    Dialog1 = "" 'Blanking out previous dialog detail
		    IF ADDR_actions = "no forwarding address provided" THEN
		        BeginDialog Dialog1, 0, 0, 186, 215, "no forwarding Address Provided"
		          CheckBox 10, 95, 165, 10, "Verif Request (DHS-2919A) Request for Contact", verif_request_checkbox
		          CheckBox 10, 105, 100, 10, "Change Report (DHS-2402)", CRF_checkbox
		          CheckBox 10, 115, 90, 10, "Shelter Form (DHS-2952)", SVF_checkbox
		          DropListBox 115, 135, 65, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", mets_addr_correspondence
		          EditBox 115, 155, 65, 15, METS_case_number
		          EditBox 50, 175, 130, 15, other_notes
		          ButtonGroup ButtonPressed
		          OkButton 75, 195, 50, 15
		          CancelButton 130, 195, 50, 15
		          GroupBox 5, 5, 175, 75, "NOTE:"
		          Text 10, 45, 160, 35, "* When a change reporting unit reports a change over the telephone or in person, the unit is NOT required to report the change on a Change Report from. "
		          GroupBox 5, 85, 175, 45, "Verification Requested:"
		          Text 10, 15, 160, 10, "Do not make any changes to STAT/ADDR."
		          Text 10, 25, 165, 15, "Do not enter a ? or unknown or other county codes on the ADDR panel."
		          Text 5, 140, 95, 10, "METS Correspondence Sent:"
		          Text 5, 160, 70, 10, "METS Case Number:"
		          Text 5, 180, 45, 10, "Other Notes:"
		        EndDialog

		        DO
		        	DO
		        		err_msg = ""
		        		DIALOG Dialog1
		        		cancel_confirmation
		        		IF verif_request_checkbox = UNCHECKED and CRF_checkbox = UNCHECKED and SVF_checkbox= UNCHECKED THEN err_msg = err_msg & vbCr & "Please elect     the verification requested and ensure forms are sent in ECF."
		        		IF mets_addr_correspondence = "YES" THEN
		        			IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid METS case number."
		        		END IF
		        		IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
		        		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		        	LOOP UNTIL err_msg = ""
		        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows ser     to password back into MAXIS
		        LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
			END IF

			return_mail_checkbox = CHECKED ' it is implied in the title of the script'
			IF ADDR_actions = "forwarding address outside MN" THEN county_code = "89 Out-of-State"
			IF ADDR_actions = "forwarding address in MN"  THEN
				new_addr_state = "MN"
				reservation_addr = "NO"
				reservation_name = "N/A"
			END IF

			Call remove_dash_from_droplist(county_list)
			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
			IF ADDR_actions = "forwarding address in MN" or ADDR_actions = "forwarding address outside MN" THEN
		        BeginDialog Dialog1, 0, 0, 206, 305, "Mail has been returned with forwarding address"
		          CheckBox 10, 15, 75, 10, "Returned Mail/Other", return_mail_checkbox
		          CheckBox 100, 15, 90, 10, "Shelter Form (DHS-2952)", SVF_checkbox
		          CheckBox 10, 25, 80, 10, "Verification Request", verif_request_checkbox
		          CheckBox 100, 25, 100, 10, "Change Report (DHS-2402)", CRF_checkbox
		          DropListBox 40, 60, 155, 15, "Select:"+chr(9)+"Own Housing: Lease, Mortgage, or Roommate"+chr(9)+"Family/Friends Due to Economic Hardship"+chr(9)+"Service Provider-Foster Care Group Home"+chr(9)+"Hospital/Treatment/Detox/Nursing Home"+chr(9)+"Jail/Prison/Juvenile Detention Center"+chr(9)+"Hotel/Motel"+chr(9)+"Emergency Shelter"+chr(9)+"Place Not Meant for Housing"+chr(9)+"Declined"+chr(9)+"Unknown", living_situation
		          EditBox 40, 80, 155, 15, new_addr_line_one
		          EditBox 40, 100, 155, 15, new_addr_line_two
		          EditBox 40, 120, 155, 15, new_addr_city
		          EditBox 40, 140, 20, 15, new_addr_state
		          EditBox 155, 140, 40, 15, new_addr_zip
		          DropListBox 55, 165, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", homeless_addr_yn
		          DropListBox 55, 185, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", reservation_addr_yn
		          DropListBox 130, 165, 65, 15, "Select:"+chr(9)+ county_list, county_code
		          DropListBox 55, 205, 140, 15, "Select:"+chr(9)+"N/A"+chr(9)+"Bois Forte-Nett Lake"+chr(9)+"Bois Forte-Vermillion Lk"+chr(9)+"Fond du Lac"+chr(9)+"Grand Portage"+chr(9)+"Leach Lake"+chr(9)+"Lower Sioux"+chr(9)+"Mille Lacs"+chr(9)+"Prairie Island Community"+chr(9)+"Red Lake"+chr(9)+"Shakopee Mdewakanton"+chr(9)+"Upper Sioux"+chr(9)+"White Earth", reservation_name
		          DropListBox 140, 230, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", mets_addr_correspondence
		          EditBox 140, 245, 55, 15, mets_case_number
		          EditBox 55, 265, 140, 15, other_notes
		          ButtonGroup ButtonPressed
		        	OkButton 85, 285, 50, 15
		        	CancelButton 145, 285, 50, 15
		          GroupBox 5, 5, 195, 35, "Verification Received:"
		          GroupBox 5, 40, 195, 185, "New Address:"
		          Text 10, 85, 25, 10, "Street:"
		          Text 20, 125, 15, 10, "City:"
		          Text 20, 145, 20, 10, "State:"
		          Text 120, 145, 30, 10, "Zip code:"
		          Text 100, 170, 25, 10, "County:"
		          Text 15, 170, 35, 10, "Homeless:"
		          Text 10, 50, 55, 10, "Living Situation:"
		          Text 10, 190, 45, 10, "Reservation:"
		          Text 10, 210, 25, 10, "Name:"
		          Text 10, 235, 95, 10, "METS correspondence sent:"
		          Text 10, 250, 70, 10, "METS case number:"
		          Text 10, 270, 40, 10, "Other notes:"
		        EndDialog

		    	DO
		    		DO
		    			err_msg = ""
		    			DIALOG Dialog1
		    			cancel_without_confirmation
		    			IF new_addr_line_one = "" THEN err_msg = err_msg & vbCr & "Please complete the street address the client in now living at."
		    			IF new_addr_city = "" THEN err_msg = err_msg & vbCr & "Please complete the city in which the client in now living."
		    			IF new_addr_state = "" THEN err_msg = err_msg & vbCr & "Please complete the state in which the client in now living."
		    			IF new_addr_zip = "" OR (new_addr_zip <> "" AND len(new_addr_zip) > 5) THEN err_msg = err_msg & vbNewLine & "Please only enter a 5 digit zip code."     'Makes sure there is a numeric zip
		    			IF homeless_addr_yn = "Select:" THEN err_msg = err_msg & vbCr & "Please advise whether the client has reported homelessness."
		    			IF living_situation = "Select:" THEN err_msg = err_msg & vbCr & "Please select the client's living situation - Unknown should not if it is avoidable."
		    			IF reservation_addr = "Select:" THEN err_msg = err_msg & vbCr & "Please select if client is living on the reservation."
		    			IF reservation_addr = "YES" THEN
		    				IF reservation_name = "Select:" THEN err_msg = err_msg & vbCr & "Please select the name of the reservation the client is living on."
		    			END IF
		    			IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
		    			IF mets_addr_correspondence = "YES" THEN
		    				IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
		    			END IF
		    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		    		LOOP UNTIL err_msg = ""
		    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to password back into MAXIS
		    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

				county_code_number = left(county_code, 2)
				county_code = right(county_code, len(county_code)-3)

		    	new_addr_line_one = trim(new_addr_line_one)
		        new_addr_line_two = trim(new_addr_line_two)
		        new_addr_city = trim(new_addr_city)
		        new_addr_state = trim(new_addr_state)
		        new_addr_zip = trim(new_addr_zip)

				IF homeless_addr_yn = "YES" THEN homeless_addr = "Y"
				IF homeless_addr_yn = "NO" THEN homeless_addr = "N"
				IF reservation_addr_yn = "YES" THEN reservation_addr = "Y"
				IF reservation_addr_yn = "NO" THEN reservation_addr = "N"

				IF living_situation = "Own Housing: Lease, Mortgage or Roommate" THEN living_situation_code = "01"
				IF living_situation = "Family/Friends Due to Economic Hardship" THEN living_situation_code = "02"
				IF living_situation = "Service Provider-Foster Care Group Home" THEN living_situation_code = "03"
				IF living_situation = "Hospital/Treatment/Detox/Nursing Home" THEN living_situation_code = "04"
				IF living_situation = "Jail/Prison/Juvenile Detention Center" THEN living_situation_code = "05"
				IF living_situation = "Hotel/Motel" THEN living_situation_code = "06"
				IF living_situation = "Emergency Shelter" THEN living_situation_code = "07"
				IF living_situation = "Place Not Meant for housing" THEN living_situation_code = "08"
				IF living_situation = "Declined" THEN living_situation_code = "09"
				IF living_situation = "Unknown" THEN living_situation_code = "10"

				IF reservation_name = "Bois Forte-Deer Creek" THEN reservation_code = "BD"
				IF reservation_name = "Bois Forte-Nett Lake" THEN reservation_code = "BN"
				IF reservation_name = "Bois Forte-Vermillion Lk" THEN reservation_code = "BV"
				IF reservation_name = "Fond du Lac" THEN reservation_code = "FL"
				IF reservation_name = "Grand Portage" THEN reservation_code = "GP"
				IF reservation_name = "Leach Lake" THEN reservation_code = "LL"
				IF reservation_name = "Lower Sioux" THEN reservation_code = "LS"
				IF reservation_name = "Mille Lacs" THEN reservation_code = "ML"
				IF reservation_name = "Prairie Island Community" THEN reservation_code = "PL"
				IF reservation_name = "Red Lake" THEN reservation_code = "RL"
				IF reservation_name = "Shakopee Mdewakanton" THEN reservation_code = "SM"
				IF reservation_name = "Upper Sioux" THEN reservation_code = "US"
				IF reservation_name = "White Earth" THEN reservation_code = "WE"

		       'Go to ADDR to update
		        Call navigate_to_MAXIS_screen("STAT", "ADDR")
				PF9
				EMWriteScreen MAXIS_footer_month, 04, 43 'to avoid error_msg = "PANEL EFFECTIVE DATE MUST EQUAL BENEFIT MONTH/YEAR" '
				EMWriteScreen "01", 04, 46 'using the first to default '
				EMWriteScreen MAXIS_footer_year, 04, 49
		        ' the reason we are clearing the residence ADDR here is because we get an inhibiting message so we have to clear the residence - residence is required mailing is not but mailing is what is used for ECF
				IF residential_address_confirmed = "YES" THEN
		    	    Call clear_line_of_text(6, 43)'Residence street'
		    	    Call clear_line_of_text(7, 43)'Residence street line two'
		    	    Call clear_line_of_text(8, 43)'Residence City'
		    	    Call clear_line_of_text(9, 43)'Residence zip'

		            EMwritescreen new_addr_line_one, 6, 43'New residence street'
		            EMwritescreen new_addr_line_two, 7, 43 ' New residence street line two'
		            EMwritescreen new_addr_city, 8, 43 'new mailing City'
		            EMwritescreen new_addr_state, 8, 66 'new mailing state'
		            EMwritescreen new_addr_zip, 9, 43
				END IF
				IF mailing_address_confirmed = "YES" THEN
		            Call clear_line_of_text(13, 43)'Mailing Street'
		            Call clear_line_of_text(14, 43)'Mailing street line two'
		            Call clear_line_of_text(15, 43)'Mailing City'
		            Call clear_line_of_text(16, 43)'Mailing Zip'

		    	    EMwritescreen new_addr_line_one, 13, 43
		            EMwritescreen new_addr_line_two, 14, 43
		            EMwritescreen new_addr_city, 15, 43 'new mailing City'
		            EMwritescreen new_addr_state, 16, 43'new mailing state'
		            EMwritescreen new_addr_zip, 16, 52
				END IF
		        IF ADDR_phone_1A = "___" THEN Call clear_line_of_text(17, 67)'removing phone code if no number is provided'
		        EMwritescreen county_code_number, 9, 66
		        EMwritescreen "OT", 9, 74
		        EMwritescreen homeless_addr, 10, 43
		        EMwritescreen reservation_addr, 10, 74 'yes no'
		        IF reservation_addr = "N" THEN Call clear_line_of_text(11, 74) 'removing the reservation name'
		        EMwritescreen reservation_code, 11, 74 'Name of Reservation'
		        EMwritescreen living_situation_code, 11, 43
			    TRANSMIT
				EMReadScreen pop_up_msg, 51, 3, 6 ' this is the message at the top of of the screen Warning: Mail to this Residence address will not be
				IF pop_up_msg = "Warning: Mail to this Residence address will not be" THEN ' MAILING ADDRESS IS STANDARDIZED'
					MsgBox "*** NOTICE!!! ***" & vbNewLine & "MAILING ADDRESS IS NOT STANDARDIZED" & vbNewLine
					TRANSMIT ' mail will not be delivered confirm'
				END IF
				TRANSMIT' confirm popup
				TRANSMIT 'confirm address
			    EMReadScreen error_msg, 75, 24, 2 ' this is the message at the bottom of the screen'
			    error_msg = TRIM(error_msg)
			    IF error_msg <> "" THEN
			    	MsgBox "*** NOTICE!!! ***" & vbNewLine & error_msg & vbNewLine
			    	IF error_msg = "WARNING: EFFECTIVE DATE HAS CHANGED - REVIEW LIVING SITUATION" THEN 'This is the only one i dont care about'
						TRANSMIT
					END IF
			    END IF
			END IF

		    IF ADDR_actions = "no response received" THEN
				CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)'

				active_programs = ""        'Creates a variable that lists all the active.
				IF ga_case = TRUE THEN active_programs = active_programs & "GA, "
				IF msa_case = TRUE THEN active_programs = active_programs & "MFIP, "
				IF mfip_case = TRUE THEN active_programs = active_programs & "MFIP, "
				IF dwp_case = TRUE THEN active_programs = active_programs & "DWP, "
				IF grh_case = TRUE THEN active_programs = active_programs & "GRH, "
				IF snap_case = TRUE THEN active_programs = active_programs & "SNAP, "
				IF ma_case = TRUE THEN active_programs = active_programs & "HC, "
				IF msp_case = TRUE THEN active_programs = active_programs & "MSP, "
				IF unknown_cash_pending = TRUE THEN active_programs = active_programs & "unknown pending, "
				'IF case_active = TRUE THEN
				'IF case_pending = TRUE THEN
				'TODO if both cash one and cash two are active the pact panel will need tobe updated manually
				'
		        active_programs = trim(active_programs)  'trims excess spaces of active_programs
		        If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

		        '-------------------------------------------------------------------------------------------------DIALOG
		        Dialog1 = "" 'Blanking out previous dialog detail
				BeginDialog Dialog1, 0, 0, 361, 90, "Client Has Not Responded to Request/No Response Received"
				  CheckBox 10, 15, 75, 10, "Returned Mail/Other", return_mail_checkbox
				  CheckBox 10, 25, 80, 10, "Verification Request", verif_request_checkbox
				  CheckBox 100, 15, 90, 10, "Shelter Form (DHS-2952)", SVF_checkbox
				  CheckBox 100, 25, 100, 10, "Change Report (DHS-2402)", CRF_checkbox
				  EditBox 310, 5, 45, 15, date_requested
				  DropListBox 310, 25, 45, 20, "Select One:"+chr(9)+"YES"+chr(9)+"NO", ECF_reveiwed
				  EditBox 55, 45, 145, 15, other_notes
				  ButtonGroup ButtonPressed
				    OkButton 260, 45, 45, 15
				    CancelButton 310, 45, 45, 15
				  Text 205, 15, 100, 10, "Date verification(s) requested:"
				  Text 5, 50, 45, 10, "Other Notes:"
				  Text 205, 30, 105, 10, "ECF reviewed for verifications?"
				  Text 5, 70, 350, 20, "Allow the household 10 days to respond before proceeding with a termination notice and ensure that the action is appropriate for the active programs. See POLI/TEMP: TE02.08.012"
				  GroupBox 5, 5, 195, 35, "Verification(s) Requested:"
				EndDialog

		        DO
		        	DO
		        		err_msg = ""
		        		DIALOG Dialog1
		        		cancel_without_confirmation
		        		If isdate(date_requested) = FALSE THEN  err_msg = err_msg & vbnewline & "Please enter the date verifications were requested."
						IF Cdate(date_requested) > cdate(date) = TRUE THEN  err_msg = err_msg & vbnewline & "You must enter an actual date that is not in the future."
		        		IF ECF_reveiwed = "Select One:" THEN  err_msg = err_msg & vbnewline & "Please review ECF to ensure the requested verifications are not on file."
		        		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		        	LOOP UNTIL err_msg = ""
		        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to      password back into MAXIS
		        LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

				IF family_cash_case = TRUE OR adult_cash_case = TRUE OR ga_case = TRUE OR msa_case = TRUE OR mfip_case = TRUE OR dwp_case = TRUE OR	grh_case = TRUE OR snap_case = TRUE THEN case_active = TRUE

		        'per POLI/TEMP this only pretains to active cash and snap '
		        IF case_active  = TRUE THEN
		        	CALL MAXIS_background_check
		        	CALL navigate_to_MAXIS_screen("STAT", "PACT")
		        	'Checking to see if the PACT panel is empty, if not it create a new panel'
		        	EMReadScreen panel_number, 1, 02, 73
		        	If panel_number = "0" then
		        		EMWriteScreen "NN", 20,79 'cursor is automatically set to 06, 58'
		        		TRANSMIT
		        		EMReadScreen MISC_error_msg,  74, 24, 02
		        		IF trim(MISC_error_msg) = "" THEN
		        			case_note_only = FALSE
		        		else
		        			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYESNO + vbQuestion, "Message handling")
		        			IF maxis_error_check = vbNO THEN
		        				case_note_only = FALSE 'this will case note only'
		        			END IF
		        			IF maxis_error_check= vbYES THEN
		        				case_note_only = TRUE 'this will update the panels and case note'
		        			END IF
		        		END IF
		        	ELSE
		        		PF9
		        		EMReadScreen open_cash1, 2, 6, 43
		        		EMReadScreen open_cash2, 2, 8, 43
		        		EMReadScreen open_grh, 2, 10, 43
		        		EMReadScreen open_snap, 2, 12, 43
		        		IF cash_active  = TRUE THEN EMWriteScreen "3", 6, 58
		        		IF cash2_active = TRUE THEN EMWriteScreen "3", 8, 58
		        		IF snap_case = TRUE THEN EMWriteScreen "4", 12, 58
		        		IF grh_active = TRUE THEN EMWriteScreen "3", 10, 58
		        		TRANSMIT
		        		EMReadScreen pop_upmsg, 7, 11, 08
		        		IF pop_upmsg = "WARNING" THEN
		        			EmWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
		        			TRANSMIT
		        		END IF
		        	END IF

		        	IF case_note_only = FALSE THEN
		        		IF case_active  = TRUE THEN EMWriteScreen "3", 6, 58
		        		'IF cash2_active = TRUE THEN EMWriteScreen "3", 8, 58 'TODO need to expalin this is the reason for reading PROG'
		        		IF SNAP_active  = TRUE THEN EMWriteScreen "4", 12, 58
		        		IF grh_active = TRUE THEN EMWriteScreen "3", 10, 58
		        		TRANSMIT
		        		EMReadScreen pop_upmsg, 7, 11, 08
		        		IF pop_upmsg = "WARNING" THEN
		        			EmWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
		        			TRANSMIT
		        		END IF
		        	END IF
		        END IF
			END IF
		END IF'confirmed address from maxis
		pending_verifs = ""
		IF return_mail_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Returned Mail/Other, "
		IF SVF_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Shelter Form (DHS-2952), "
		IF verif_request_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Verification Request (DHS-2919A) Request for Contact, "
		IF CRF_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Change Report (DHS-2402), "
		'-------------------------------------------------------------------trims excess spaces of pending_verifs
		pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs
		IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
		'----------------------------------------------------------------------------------------------------TIKL
		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
		IF ADDR_actions <> "no response recieved" THEN Call create_TIKL("Returned mail rec'd contact from the client should have occurred regarding address change. If no response-verbal or written, please take appropriate action.", 10, date, True, TIKL_note_text)
		IF mailing_addr_line_two <> "" THEN mailing_addr_line_two = mailing_addr_line_two & " "
		IF resi_addr_line_two <> "" THEN resi_addr_line_two = resi_addr_line_two & " "
		'starts a blank case note
		CALL start_a_blank_case_note
		CALL write_variable_in_CASE_NOTE("Returned mail received " & ADDR_actions)
		CALL write_bullet_and_variable_in_CASE_NOTE("Received on", date_received)
		 'Address Detail
		 IF mailing_address_confirmed = "YES" THEN
		 	Call write_variable_in_CASE_NOTE("* The address on ADDR was reviewed and is correct.")
			CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & mail_line_one)
			CALL write_variable_in_CASE_NOTE("                             " & mail_line_two & mail_city_line & ", " & mail_state_line & " " &   mail_zip_line)
		 ELSE
			CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & resi_addr_line_one)
			CALL write_variable_in_CASE_NOTE("                               " & resi_addr_line_two & resi_addr_city & ", " & resi_addr_state & " " & resi_addr_zip)
		END IF
		IF homeless_addr_yn = "YES" Then Call write_variable_in_CASE_NOTE("* Household is homeless")
		Call write_variable_in_CASE_NOTE("* Mail received reports living in: " & county_code & " County and CM 0008.06.21 - was reviewed if applicable")
		IF reservation_addr = "YES" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
		Call write_bullet_and_variable_in_CASE_NOTE("Living Situation", living_situation)
		Call write_bullet_and_variable_in_CASE_NOTE("Address Detail", notes_on_address)
		IF ADDR_actions <> "no response recieved"  THEN
			CALL write_bullet_and_variable_in_case_note("Verification(s) Received", pending_verifs)
			IF mailing_address_confirmed = "YES" THEN
				CALL write_variable_in_CASE_NOTE("* Mailing address updated:  " & new_addr_line_one)
		    	CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
			END IF
		 	IF residential_address_confirmed = "YES" THEN
				CALL write_variable_in_CASE_NOTE("* Residential address updated:  " & new_addr_line_one)
				CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
			END IF
		ELSE
			CALL write_bullet_and_variable_in_case_note("Verification(s) requested", pending_verifs)
			CALL write_bullet_and_variable_in_case_note("Verification(s) request date", date_requested)
			IF ECF_review_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE ("* ECF reviewed for requested verifications")
			CALL write_variable_in_CASE_NOTE ("* PACT panel entered per POLI/TEMP TE02.13.10")
		END IF
		IF mets_addr_correspondence <> "Select:" THEN CALL write_bullet_and_variable_in_CASE_NOTE("METS correspondence sent", mets_addr_correspondence)
		CALL write_bullet_and_variable_in_CASE_NOTE("METS case number", METS_case_number)
		CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
		CALL write_variable_in_CASE_NOTE ("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)

		'Checks if this is a MNsure case and pops up a message box with instructions if the ADDR is incorrect.
		IF METS_case_number <> "" and mets_addr_correspondence = "NO" THEN MsgBox "Please update the METS ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. Change In Circumstance Process)"

		IF ADDR_actions <> "no response recieved" THEN
			script_end_procedure_with_error_report("Success! TIKL has been set for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over he telephone or in person, the unit is not required to also report the change on a Change Report from. ")
		ELSE
			script_end_procedure_with_error_report("Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter a worker comment in PEC/WCOM.")
		END IF

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
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("02/13/2020", "Updated the zip code to only allow for 5 characters.", "MiKayla Handley, Hennepin County")
call changelog_update("06/06/2019", "Initial version. Rewitten per POLI/TEMP.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT------------------------------------------------------------------------------------------------------
'Connects to BLUEZONE
EMConnect ""
'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'Running the initial dialog
BeginDialog Dialog1, 0, 0, 211, 85, "Returned Mail"
  EditBox 55, 5, 40, 15, maxis_case_number
  EditBox 165, 5, 40, 15, date_received
  DropListBox 90, 25, 115, 15, "Select:"+chr(9)+"Forwarding address in MN"+chr(9)+"Forwarding address outside MN"+chr(9)+"No forwarding address provided"+chr(9)+"Client has not responded to request", ADDR_actions
  EditBox 90, 45, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 110, 65, 45, 15
    CancelButton 160, 65, 45, 15
  Text 5, 30, 85, 10, "Mail Has Been Returned:"
  Text 5, 50, 60, 10, "Worker Signature:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 105, 10, 50, 10, "Date Received:"
EndDialog

DO
    DO
    	err_msg = ""
    	DIALOG Dialog1
    	cancel_confirmation
    	IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		If isdate(date_received) = FALSE THEN
			err_msg = err_msg & vbnewline & "* You must enter an actual date that is not in the future and is in the footer month that you are working in."
		ElseIf Cdate(date_received) > cdate(date) = TRUE THEN
			err_msg = err_msg & vbnewline & "* You must enter an actual date that is not in the future and is in the footer month that you are working in."
		End If
		IF ADDR_actions = "Select:" THEN err_msg = err_msg & vbCr & "* Please chose an action for the returned mail."
    	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'MAXIS_footer_month = right("00" & DatePart("m", date_received), 2)
MAXIS_footer_month = DatePart("m", date_received)
MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
MAXIS_footer_year = DatePart("yyyy", date_received)
MAXIS_footer_year = right(MAXIS_footer_year, 2)

CALL back_to_SELF

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'EMReadScreen err_msg, 7, 24, 02
'IF err_msg = "BENEFIT" THEN	script_end_procedure_with_error_report ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

'Reading the program status
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen snap_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
'----------------------------------------------------------------------------------------------------ACTIVE program coding
EMReadScreen cash1_prog_check, 2, 6, 67     'Reading cash 1
EMReadScreen cash2_prog_check, 2, 7, 67     'Reading cash 2 this is funky if inac they program will show up on pact but not on prog
EMReadScreen emer_prog_check, 2, 8, 67      'EMER Program

'Logic to determine if MFIP is active
IF cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "MS" THEN
	IF cash1_status_check = "ACTV" THEN cash_active = TRUE
END IF
IF cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "MS" THEN
	IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
END IF
IF emer_prog_check = "EG" and emer_status_check = "ACTV" THEN emer_active = TRUE
IF emer_prog_check = "EA" and emer_status_check = "ACTV" THEN emer_active = TRUE

IF cash1_status_check = "ACTV" THEN cash_active  = TRUE
IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
IF snap_status_check  = "ACTV" THEN SNAP_active  = TRUE
IF grh_status_check   = "ACTV" THEN grh_active   = TRUE
IF ive_status_check   = "ACTV" THEN IVE_active   = TRUE
IF hc_status_check    = "ACTV" THEN hc_active    = TRUE
IF cca_status_check   = "ACTV" THEN cca_active   = TRUE

active_programs = ""        'Creates a variable that lists all the active.
IF cash_active = TRUE or cash2_active = TRUE THEN active_programs = active_programs & "CASH, "
IF emer_active = TRUE THEN active_programs = active_programs & "Emergency, "
IF grh_active  = TRUE THEN active_programs = active_programs & "GRH, "
IF snap_active = TRUE THEN active_programs = active_programs & "SNAP, "
IF ive_active  = TRUE THEN active_programs = active_programs & "IV-E, "
IF hc_active   = TRUE THEN active_programs = active_programs & "HC, "
IF cca_active  = TRUE THEN active_programs = active_programs & "CCA"

active_programs = trim(active_programs)  'trims excess spaces of active_programs
If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

CALL navigate_to_MAXIS_screen("STAT", "ADDR")
'EMReadScreen MAXIS_footer_month, 2, 20, 55
'EMReadScreen MAXIS_footer_year, 2, 20, 58
EMreadscreen resi_addr_line_one, 22, 6, 43
resi_addr_line_one = replace(resi_addr_line_one, "_", "")
EMreadscreen resi_addr_line_two, 22, 7, 43
resi_addr_line_two = replace(resi_addr_line_two, "_", "")
EMreadscreen resi_addr_city, 15, 8, 43
resi_addr_city = replace(resi_addr_city, "_", "")
EMreadscreen resi_addr_state, 2, 8, 66
resi_addr_state = replace(resi_addr_state, "_", "")
EMreadscreen resi_addr_zip, 7, 9, 43
resi_addr_zip = replace(resi_addr_zip, "_", "")
EMreadscreen addr_county, 2, 9, 66
EMreadscreen addr_verif, 2, 9, 74
EMreadscreen addr_homeless, 1, 10, 43
'EMreadscreen ADDR_reservation, 1, 10, 74
EMreadscreen mailing_addr_line_one, 22, 13, 43
mailing_addr_line_one = replace(mailing_addr_line_one, "_", "")
EMreadscreen mailing_addr_line_two, 22, 14, 43
mailing_addr_line_two = replace(mailing_addr_line_two, "_", "")
EMreadscreen mailing_addr_city, 15, 15, 43
mailing_addr_city = replace(mailing_addr_city, "_", "")
EmReadScreen mailing_addr_state, 2, 16, 43
mailing_addr_state = replace(mailing_addr_state, "_", "")
EMreadscreen mailing_addr_zip, 7, 16, 52
mailing_addr_zip = replace(mailing_addr_zip, "_", "")
EMreadscreen ADDR_phone_1A, 3, 17, 45						'Has to split phone numbers up into three parts each
EMreadscreen ADDR_phone_2B, 3, 17, 51
EMreadscreen ADDR_phone_3C, 4, 17, 55
EMreadscreen ADDR_phone_2A, 3, 18, 45
EMreadscreen ADDR_phone_2B, 3, 18, 51
EMreadscreen ADDR_phone_2C, 4, 18, 55
EMreadscreen ADDR_phone_3A, 3, 19, 45
EMreadscreen ADDR_phone_3B, 3, 19, 51
EMreadscreen ADDR_phone_3C, 4, 19, 55
IF mailing_addr_line_one <> "" THEN
	maxis_addr = mailing_addr_line_one & " " & mailing_addr_line_two & " " & mailing_addr_city & " " & mailing_addr_state & " " & mailing_addr_zip
ELSE
    maxis_addr = resi_addr_line_one & " " & resi_addr_line_two & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip
END IF
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
IF ADDR_actions = "No forwarding address provided" THEN
    BeginDialog Dialog1, 0, 0, 186, 215, "No forwarding address provided"
      CheckBox 10, 95, 165, 10, "Verif Request (DHS-2919A)-Request for Contact ", verifA_sent_checkbox
      CheckBox 10, 105, 100, 10, "Change Report (DHS-2402)", CRF_sent_checkbox
      CheckBox 10, 115, 70, 10, "SVF (DHS-2952)", SHEL_form_sent_checkbox
      DropListBox 110, 135, 65, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", mets_addr_correspondence
      EditBox 110, 155, 65, 15, METS_case_number
      EditBox 50, 175, 130, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 75, 195, 50, 15
        CancelButton 130, 195, 50, 15
      GroupBox 5, 5, 175, 75, "NOTE:"
      Text 10, 15, 160, 10, "Do not make any changes to STAT/ADDR."
      Text 10, 25, 165, 15, "Do not enter a ? or unknown or other county codes on the ADDR panel."
      Text 10, 45, 160, 35, "* When a change reporting unit reports a change over the telephone or in person, the unit is not required to also report the change on a Change Report from. "
      GroupBox 5, 85, 175, 45, "Verification Requested:"
      Text 5, 140, 95, 10, "METS Correspondence Sent:"
      Text 5, 160, 70, 10, "METS Case Number:"
      Text 5, 180, 45, 10, "Other Notes:"
    EndDialog

	DO
	    DO
	    	err_msg = ""
	    	DIALOG Dialog1
	    	cancel_without_confirmation
			IF verifA_sent_checkbox = UNCHECKED and CRF_sent_checkbox = UNCHECKED and SHEL_form_sent_checkbox= UNCHECKED THEN err_msg = err_msg & vbCr & "Please select the verifcation requested and ensure forms are sent in ECF."
	    	IF mets_addr_correspondence = "YES" THEN
				IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
			END IF
			IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
	    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    LOOP UNTIL err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
END IF

IF ADDR_actions = "Forwarding address outside MN" THEN county_code = "89 Out-of-State"
IF ADDR_actions = "Forwarding address in MN"  then new_addr_state = "MN"
Call remove_dash_from_droplist(county_list)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
IF ADDR_actions = "Forwarding address in MN" or ADDR_actions = "Forwarding address outside MN" THEN
    BeginDialog Dialog1, 0, 0, 206, 305, "Mail has been returned with forwarding address"
      Text 10, 15, 185, 10, maxis_addr
	  CheckBox 10, 35, 100, 10, "Verif Request (DHS-2919A)", verifA_sent_checkbox
   	  CheckBox 120, 35, 70, 10, "SVF (DHS-2952)", SHEL_form_sent_checkbox
      CheckBox 10, 45, 100, 10, "Change Report (DHS-2402)", CRF_sent_checkbox
      CheckBox 120, 45, 75, 15, "Returned Mail/Other", return_mail_checkbox
      DropListBox 45, 75, 150, 15, "Select:"+chr(9)+"Own Housing: Lease, Mortgage, or Roommate"+chr(9)+"Family/Friends Due to Economic Hardship"+chr(9)+"Service Provider-Foster Care Group Home"+chr(9)+"Hospital/Treatment/Detox/Nursing Home"+chr(9)+"Jail/Prison/Juvenile Detention Center"+chr(9)+"Hotel/Motel"+chr(9)+"Emergency Shelter"+chr(9)+"Place Not Meant for housing"+chr(9)+"Declined"+chr(9)+"Unknown", living_situation
      EditBox 40, 95, 155, 15, new_addr_line_one
      EditBox 40, 115, 155, 15, new_addr_line_two
      EditBox 40, 135, 155, 15, new_addr_city
      EditBox 40, 155, 20, 15, new_addr_state
      EditBox 155, 155, 40, 15, new_addr_zip
      DropListBox 55, 175, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", homeless_addr
	  DropListBox 130, 175, 65, 15, "Select:"+chr(9)+county_list, county_code
      DropListBox 55, 190, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", reservation_addr
      DropListBox 55, 205, 140, 15, "Select:"+chr(9)+"N/A"+chr(9)+"Bois Forte-Nett Lake"+chr(9)+"Bois Forte-Vermillion Lk"+chr(9)+"Fond du Lac"+chr(9)+"Grand Portage"+chr(9)+"Leach Lake"+chr(9)+"Lower Sioux"+chr(9)+"Mille Lacs"+chr(9)+"Prairie Island Community"+chr(9)+"Red Lake"+chr(9)+"Shakopee Mdewakanton"+chr(9)+"Upper Sioux"+chr(9)+"White Earth", reservation_name
      DropListBox 140, 230, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", mets_addr_correspondence
      EditBox 140, 245, 55, 15, METS_case_number
      EditBox 55, 265, 140, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 85, 285, 50, 15
        CancelButton 145, 285, 50, 15
      GroupBox 5, 5, 195, 20, "Address in MAXIS:"
      GroupBox 5, 25, 195, 35, "Verification Received:"
      GroupBox 5, 60, 195, 165, "New Address:"
      Text 10, 100, 25, 10, "Street:"
      Text 20, 140, 15, 10, "City:"
      Text 20, 160, 20, 10, "State:"
      Text 120, 160, 30, 10, "Zip code:"
      Text 100, 180, 25, 10, "County:"
      Text 10, 180, 35, 10, "Homeless:"
      Text 10, 80, 35, 10, "Living Sit:"
      Text 10, 195, 45, 10, "Reservation:"
      Text 10, 210, 25, 10, "Name:"
      Text 10, 235, 95, 10, "METS correspondence sent:"
      Text 10, 250, 70, 10, "METS case number:"
      Text 10, 270, 40, 10, "Other notes:"
    EndDialog

    DO
    	DO
    		err_msg = ""
    		DIALOG Dialog1
    		cancel_without_confirmation
			IF new_addr_line_one = "" THEN err_msg = err_msg & vbCr & "Please complete the street address the client in now living at."
			IF new_addr_city = "" THEN err_msg = err_msg & vbCr & "Please complete the city in which the client in now living."
			IF new_addr_state = "" THEN err_msg = err_msg & vbCr & "Please complete the state in which the client in now living."
			IF new_addr_zip = "" OR (new_addr_zip <> "" AND len(new_addr_zip) > 5) THEN err_msg = err_msg & vbNewLine & "*Please only enter a 5 digit zip code." 'Makes sure there is a numeric zip
			IF homeless_addr = "Select:" THEN err_msg = err_msg & vbCr & "Please advise whether the client has reported homlessness."
			IF living_situation = "Select:" THEN err_msg = err_msg & vbCr & "Please select the client's living situation - Unknown should not be selected as this should be covered in the interview."
			IF reservation_addr = "Select:" THEN err_msg = err_msg & vbCr & "Please select if client is living on the reservation."
			IF reservation_addr = "YES" THEN
				IF reservation_name = "Select:" THEN err_msg = err_msg & vbCr & "Please select the name of the reservation the client is living on."
			END IF
      		IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
			IF mets_addr_correspondence = "YES" THEN
				IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
			END IF
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	county_code_number = left(county_code, 2)
	county_code = right(county_code, len(county_code)-3)
END IF

IF homeless_addr = "YES" THEN homeless_addr_code = "Y"
IF homeless_addr = "NO" THEN homeless_addr_code = "N"
IF reservation_addr = "YES" THEN reservation_addr_code = "Y"
IF reservation_addr = "NO" THEN reservation_addr_code = "N"

IF living_situation = "Own Housing: Lease, Mortgage or Roommate" THEN living_situation_code = "01"
IF living_situation = "Family/Friends Due to Economic Hardship" THEN living_situation_code = "02"
IF living_situation = "Service Provider-Foster Care Group Home" THEN living_situation_code = "03"
IF living_situation = "Hospital/Treatment/Detox/Nursing Home" THEN living_situation_code = "04"
IF living_situation = "Jail/Prison/Juvenile Detention Center" THEN living_situation_code = "05"
IF living_situation = "Hotel/Motel" THEN living_situation_code = "06"
IF living_situation = "Emergency Shelter" THEN living_situation_code = "07"
IF living_situation = "Place Not Meant for housing" THEN living_situation_code = "08"
IF living_situation = "Declined" THEN living_situation_code = "09"
IF living_situation = "Unknown" THEN living_situation_code = "10"

IF reservation_name = "Bois Forte-Deer Creek" THEN rez_code = "BD"
IF reservation_name = "Bois Forte-Nett Lake" THEN rez_code = "BN"
IF reservation_name = "Bois Forte-Vermillion Lk" THEN rez_code = "BV"
IF reservation_name = "Fond du Lac" THEN rez_code = "FL"
IF reservation_name = "Grand Portage" THEN rez_code = "GP"
IF reservation_name = "Leach Lake" THEN rez_code = "LL"
IF reservation_name = "Lower Sioux" THEN rez_code = "LS"
IF reservation_name = "Mille Lacs" THEN rez_code = "ML"
IF reservation_name = "Prairie Island Community" THEN rez_code = "PL"
IF reservation_name = "Red Lake" THEN rez_code = "RL"
IF reservation_name = "Shakopee Mdewakanton" THEN rez_code = "SM"
IF reservation_name = "Upper Sioux" THEN rez_code = "US"
IF reservation_name = "White Earth" THEN rez_code = "WE"

new_addr_line_one = trim(new_addr_line_one)
new_addr_line_two = trim(new_addr_line_two)
new_addr_city = trim(new_addr_city)
new_addr_state = trim(new_addr_state)
new_addr_zip = trim(new_addr_zip)
new_addr_line_one = UCASE(new_addr_line_one)
new_addr_line_two = UCASE(new_addr_line_two)
new_addr_city = UCASE(new_addr_city)
new_addr_state = UCASE(new_addr_state)
county_code = UCASE(county_code)

IF ADDR_actions = "Forwarding address in MN" or ADDR_actions = "Forwarding address outside MN" THEN
	PF9
	EmWriteScreen "01", 04, 46
	EmWriteScreen MAXIS_footer_month, 04, 43
	EmWriteScreen MAXIS_footer_year, 04, 49

	EMReadScreen error_check, 2, 24, 2	'making sure we can actually update this case.
	error_check = trim(error_check)
	If error_check <> "" then script_end_procedure("Unable to update this case. Please review case, and run the script again if applicable.")

	IF ADDR_actions = "Forwarding address outside MN" THEN' the reason we are changing ADDR here is because we get an inhibiting message  COUNTY OF RESIDENCE MUST BE 89 WHEN STATE IS NOT MN
	    Call clear_line_of_text(6, 43)'Residence street'
	    Call clear_line_of_text(7, 43)'Residence street line two'
	    Call clear_line_of_text(8, 43)'Residence City'
	    Call clear_line_of_text(9, 43)'Residence zip'
		EMwritescreen new_addr_line_one, 6, 43
		EMwritescreen new_addr_line_two, 7, 43
		IF new_addr_line_two = "" THEN Call clear_line_of_text(7, 43)
	    EMwritescreen new_addr_city, 8, 43 'Residence City'
	    EMwritescreen new_addr_state, 8, 66	'Defaults to MN for all cases at this time
	    EMwritescreen new_addr_zip, 9, 43
	END IF

	IF ADDR_actions = "Forwarding address in MN" and ADDR_county = "89" THEN' the reason we are changing ADDR here is because we get an inhibiting message  COUNTY OF RESIDENCE MUST BE 89 WHEN STATE IS NOT MN
		Call clear_line_of_text(6, 43)'Residence street'
		Call clear_line_of_text(7, 43)'Residence street line two'
		Call clear_line_of_text(8, 43)'Residence City'
		Call clear_line_of_text(9, 43)'Residence zip'
		EMwritescreen new_addr_line_one, 6, 43
		EMwritescreen new_addr_line_two, 7, 43
		IF new_addr_line_two = "" THEN Call clear_line_of_text(7, 43)
		EMwritescreen new_addr_city, 8, 43 'Residence City'
		EMwritescreen new_addr_state, 8, 66	'Defaults to MN for all cases at this time
		EMwritescreen new_addr_zip, 9, 43
	END IF

	Call clear_line_of_text(13, 43)'Mailing Street'
	Call clear_line_of_text(14, 43)'Mailing street line two'
	Call clear_line_of_text(15, 43)'Mailing City'
	Call clear_line_of_text(16, 43)'Mailing Zip'
	IF ADDR_phone_1A = "___" THEN Call clear_line_of_text(17, 67)'removing phone code if no number is provided'
	EMwritescreen county_code_number, 9, 66
	EMwritescreen "OT", 9, 74
	EMwritescreen homeless_addr_code, 10, 43
	EMwritescreen reservation_addr_code, 10, 74
	IF reservation_addr = "NO" THEN Call clear_line_of_text(11, 74) 'removing the reseervation name'
	EMwritescreen rez_code, 11, 74
	EMwritescreen living_situation_code, 11, 43
	EMwritescreen new_addr_line_one, 13, 43 'Mailing Street'
	EMwritescreen new_addr_line_two, 14, 43 'Mailing street line two'
	IF new_addr_line_two = "" THEN Call clear_line_of_text(14, 43)
	EMwritescreen new_addr_city, 15, 43 'Mailing City'
	EMwritescreen new_addr_state, 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
	EMwritescreen new_addr_zip, 16, 52
	'Error messages'
	'PANEL EFFECTIVE DATE MUST EQUAL BENEFIT MONTH/YEAR
	'WARNING: EFFECTIVE DATE HAS CHANGED - REVIEW LIVING SITUATION
	'TYPE NOT ALLOWED WHEN PHONE ONE IS MISSING'
	'NAME OF RESERVATION IS MISSING'
	'COUNTY OF RESIDENCE MUST BE 89 WHEN STATE IS NOT MN' if the clien was previously out of state this will cause issue
	'Errors on the PACT'
	'CASH II IS INACTIVE'
	'CASE IS PENDING, USE '1' OR '3' TO DENY '

	' *** WARNING ***  1, 26'
	'RESIDENCE ADDRESS IS STANDARDIZED  3, 6
	TRANSMIT
'TODO - sending a PF11 for clarification on living situation 02-13-2020
	EMReadScreen error_msg, 75, 24, 2
	error_msg = TRIM(error_msg)
	IF error_msg <> "" THEN
		MsgBox "*** NOTICE!!! ***" & vbNewLine & error_msg & vbNewLine
		IF error_msg = "WARNING: EFFECTIVE DATE HAS CHANGED - REVIEW LIVING SITUATION" THEN
			TRANSMIT
		END IF
	END IF
END IF
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
IF ADDR_actions = "Client has not responded to request" THEN
    BeginDialog Dialog1, 0, 0, 151, 100, "Client has not responded to request"
      EditBox 100, 5, 45, 15, date_requested
      CheckBox 20, 25, 110, 10, "ECF reviewed for verifications", ECF_review_checkbox
      EditBox 50, 60, 95, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 60, 80, 40, 15
        CancelButton 105, 80, 40, 15
      Text 5, 40, 145, 20, "Allow the household 10 days to respond before proceeding with a termination notice."
      Text 5, 10, 75, 10, "Date verif requested:"
      Text 5, 65, 45, 10, "Other Notes:"
    EndDialog

    DO
    	DO
    		err_msg = ""
    		DIALOG Dialog1
    		cancel_without_confirmation
    		If isdate(date_requested) = FALSE and Cdate(date_requested) > cdate(date) = TRUE THEN  err_msg = err_msg & vbnewline & "* Please enter the date verifcations were requested."
    		IF ECF_review_checkbox <> CHECKED THEN  err_msg = err_msg & vbnewline & "* Please review ECF to ensure the requested verifications are not on file."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'per POLI/TEMP this only pretains to active cash and snap '
	IF cash_active  = TRUE or cash2_active = TRUE or SNAP_active  = TRUE THEN
		CALL MAXIS_background_check
    	CALL navigate_to_MAXIS_screen("STAT", "PACT")
    	'Checking to see if the PACT panel is empty, if not it create a new panel'
		EmReadScreen panel_number, 1, 02, 73
        If panel_number = "0" then
        	EMWriteScreen "NN", 20,79 'cursor is automatically set to 06, 58'
        	TRANSMIT
    		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
    		EmReadScreen MISC_error_msg,  74, 24, 02
    		IF trim(MISC_error_msg) = "" THEN
    	        case_note_only = FALSE
    		else
    			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYESNO + vbQuestion, "Message handling")
    			IF maxis_error_check = vbNO THEN
    				case_note_only = FALSE 'this will case note only'
    			END IF
    			IF maxis_error_check= vbYES THEN
    				case_note_only = TRUE 'this will update the panels and case note'
    			END IF
    		END IF
		ELSE
			PF9
		    EmReadScreen open_cash1, 2, 6, 43
		    EmReadScreen open_cash2, 2, 8, 43
		    EmReadScreen open_grh, 2, 10, 43
		    EmReadScreen open_snap, 2, 12, 43
			IF cash_active  = TRUE THEN EMWriteScreen "3", 6, 58
			IF cash2_active = TRUE THEN EMWriteScreen "3", 8, 58
			IF SNAP_active  = TRUE THEN EMWriteScreen "4", 12, 58
			IF grh_active = TRUE THEN EMWriteScreen "3", 10, 58
			TRANSMIT
			EMReadScreen pop_upmsg, 7, 11, 08
			IF pop_upmsg = "WARNING" THEN
				EmWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
				TRANSMIT
			END IF
		END IF

		IF case_note_only = FALSE THEN
		    IF cash_active  = TRUE THEN EMWriteScreen "3", 6, 58
		    IF cash2_active = TRUE THEN EMWriteScreen "3", 8, 58
		    IF SNAP_active  = TRUE THEN EMWriteScreen "4", 12, 58
		    IF grh_active = TRUE THEN EMWriteScreen "3", 10, 58
		    TRANSMIT
			EMReadScreen pop_upmsg, 7, 11, 08
			IF pop_upmsg = "WARNING" THEN
				EmWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
				TRANSMIT
			END IF
       	END IF
    	'END IF
    END IF
END IF

pending_verifs = ""
IF verifA_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Verification Request, "
IF SHEL_form_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "SVF, "
IF CRF_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Change Request Form, "
IF return_mail_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Returned Mail/Other, "
'-------------------------------------------------------------------trims excess spaces of pending_verifs
pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script

Call MAXIS_background_check 'checking to make sure case is out of background

'----------------------------------------------------------------------------------------------------TIKL
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF ADDR_actions <> "Client has not responded to request" THEN Call create_TIKL("Returned mail rec'd contact from the client should have occurred regarding address change. If no response-verbal or written, please take appropriate action.", 10, date, True, TIKL_note_text)

'starts a blank case note
call start_a_blank_case_note
call write_variable_in_CASE_NOTE("Returned mail rcv'd-" & ADDR_actions & " for " & active_programs & "")
IF ADDR_actions <> "Client has not responded to request" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Received on", date_received)
IF ADDR_actions = "Forwarding address in MN" or ADDR_actions = "Forwarding address outside MN" THEN
    CALL write_bullet_and_variable_in_CASE_NOTE("Living situation", living_situation)
    CALL write_bullet_and_variable_in_CASE_NOTE("Homeless", homeless_addr)
	CALL write_bullet_and_variable_in_case_note("Verification(s) received", pending_verifs)
    CALL write_variable_in_CASE_NOTE("* Mailing address updated:  " & new_addr_line_one)
    CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
    CALL write_variable_in_CASE_NOTE("                            " & county_code & " COUNTY.")
	IF reservation_addr = "YES" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
ELSE
	CALL write_bullet_and_variable_in_case_note("Verification(s) requested", pending_verifs)
END IF
CALL write_variable_in_CASE_NOTE ("---")

IF mailing_addr_line_one <> "" THEN
	CALL write_variable_in_CASE_NOTE("* Previous mailing address: " & mailing_addr_line_one)
	CALL write_variable_in_CASE_NOTE("                           " & mailing_addr_line_two & " " & mailing_addr_city & " " & mailing_addr_state & " " & mailing_addr_zip)
ELSE
	CALL write_variable_in_CASE_NOTE("* Previous residential address: " & resi_addr_line_one)
	CALL write_variable_in_CASE_NOTE("                               " & resi_addr_line_two & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip)
END IF
IF mets_addr_correspondence <> "Select:" THEN CALL write_bullet_and_variable_in_CASE_NOTE("METS correspondence sent", mets_addr_correspondence)
CALL write_bullet_and_variable_in_CASE_NOTE("METS case number", METS_case_number)
IF ADDR_actions = "Client has not responded to request" THEN
	CALL write_bullet_and_variable_in_case_note("Verification(s) request date", date_requested)
	IF ECF_review_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE ("* ECF reviewed for requested verifications")
	CALL write_variable_in_CASE_NOTE ("* PACT panel entered per POLI/TEMP TE02.13.10")
END IF
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'Checks if this is a MNsure case and pops up a message box with instructions if the ADDR is incorrect.
IF METS_case_number <> "" and mets_addr_correspondence = "NO" THEN MsgBox "Please update the MNsure ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. Change In Circumstance)"

IF ADDR_actions <> "Client has not responded to request" THEN script_end_procedure_with_error_report("Success! TIKL has been set for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over the telephone or in person, the unit is not required to also report the change on a Change Report from. ")
IF ADDR_actions = "Client has not responded to request" THEN script_end_procedure_with_error_report("Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter a worker comment in SPEC/WCOM.")
