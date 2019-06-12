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
call changelog_update("06/06/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog case_number_dlg, 0, 0, 251, 85, "Returned Mail"
  EditBox 55, 5, 40, 15, maxis_case_number
  EditBox 205, 5, 40, 15, date_received
  DropListBox 55, 25, 190, 15, "Select One:"+chr(9)+"Mail has been returned NO forwarding address"+chr(9)+"Mail has been returned with forwarding address in MN"+chr(9)+"Mail has been returned with forwarding address outside MN"+chr(9)+"Client has not responded to request for verif", ADDR_actions
  EditBox 130, 45, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 145, 65, 50, 15
    CancelButton 195, 65, 50, 15
  Text 5, 30, 45, 10, "Select action:"
  Text 65, 50, 60, 10, "Worker signature:"
  Text 5, 10, 45, 10, "Case number:"
  Text 150, 10, 50, 10, "Date received:"
EndDialog


DO
    DO
    	err_msg = ""
    	DIALOG case_number_dlg
    		IF ButtonPressed = 0 THEN stopscript
    		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
			If isdate(date_received) = FALSE then err_msg = err_msg & vbnewline & "* Please enter a date (--/--/--) in the footer month that you are working in."
			IF ADDR_actions = "Select One:" THEN err_msg = err_msg & vbCr & "Please chose an action for the returned mail."
    		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
CALL check_for_MAXIS(False)



CALL navigate_to_MAXIS_screen("STAT", "ADDR")
'Writes spreadsheet info to ADDR
EMreadscreen ADDR_line_one, 20, 6, 43
EMreadscreen ADDR_line_two, 20, 7, 43
EMreadscreen ADDR_city, 15, 8, 43
EMreadscreen ADDR_state, 2, 8, 66
EMreadscreen ADDR_zip, 5, 9, 43
EMreadscreen ADDR_county, 2, 9, 66
EMreadscreen ADDR_addr_verif, 2, 9, 74
EMreadscreen ADDR_homeless, 1, 10, 43
'EMreadscreen ADDR_reservation, 1, 10, 74
EMreadscreen ADDR_mailing_addr_line_one, 20, 13, 43
EMreadscreen ADDR_mailing_addr_line_two, 20, 14, 43
EMreadscreen ADDR_mailing_addr_city, 15, 15, 43
EMreadscreen 2, 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
EMreadscreen ADDR_mailing_addr_zip, 5, 16, 52
EMreadscreen ADDR_phone_1A, 3, 17, 45						'Has to split phone numbers up into three parts each
EMreadscreen ADDR_phone_2B, 3, 17, 51
EMreadscreen ADDR_phone_3C, 4, 17, 55
EMreadscreen ADDR_phone_2A, 3, 18, 45
EMreadscreen ADDR_phone_2B, 3, 18, 51
EMreadscreen ADDR_phone_2C, 4, 18, 55
EMreadscreen ADDR_phone_3A, 3, 19, 45
EMreadscreen ADDR_phone_3B, 3, 19, 51
EMreadscreen ADDR_phone_3C, 4, 19, 55


maxis_addr = ADDR_line_one & " " & ADDR_line_two & " " & ADDR_city & " " & ADDR_state & " " & ADDR_zip

IF ADDR_actions = "Mail has been returned NO forwarding address" THEN
    BeginDialog RETURNED_MAIL, 0, 0, 181, 185, "Mail has been returned NO forwarding address"
    CheckBox 10, 70, 70, 10, "Sent DHS-2919A", verifA_sent_checkbox
    CheckBox 85, 70, 65, 10, "Sent DHS-2952", SHEL_form_sent_checkbox
    CheckBox 10, 85, 65, 10, "Sent DHS-2402", CRF_sent_checkbox
    EditBox 110, 105, 65, 15, METS_case_number
    DropListBox 110, 125, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", MNsure_ADDR
    EditBox 50, 145, 125, 15, other_notes
    ButtonGroup ButtonPressed
    OkButton 70, 165, 50, 15
    CancelButton 125, 165, 50, 15
    Text 5, 130, 95, 10, "METS correspondence sent:"
    Text 5, 150, 40, 10, "Other notes:"
    GroupBox 5, 5, 170, 50, "NOTE:"
    GroupBox 5, 60, 170, 40, "Verification Request Form"
    Text 5, 110, 70, 10, "METS case number:"
    Text 10, 15, 160, 35, "Do not make any changes to STAT/ADDR.  Do NOT enter a ? or unknown or other county codes on the ADDR panel.  The ADDR panel is used to mail notices; the post office requires an address. "
    EndDialog

	DO
	    DO
	    	err_msg = ""
	    	DIALOG case_number_dlg
	    		IF ButtonPressed = 0 THEN stopscript
	    		IF MNsure_ADDR = "YES" THEN
					IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
				END IF
				IF MNsure_ADDR = "Select One:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
	    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    	LOOP UNTIL err_msg = ""
	    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	CALL check_for_MAXIS(False)
END IF

IF ADDR_actions = "Mail has been returned with forwarding address in MN" or ADDR_actions = "Mail has been returned with forwarding address outside MN" THEN
	BeginDialog returned_mail_update_addr, 0, 0, 201, 280, "Mail has been returned with forwarding address"
	  Text 10, 15, 180, 35, maxis_addr
      CheckBox 10, 65, 50, 10, "DHS-2919A", verifA_sent_checkbox
      CheckBox 70, 65, 45, 10, "DHS-2952", SHEL_form_sent_checkbox
      CheckBox 125, 65, 45, 10, "DHS-2402", CRF_sent_checkbox
      DropListBox 135, 90, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", update_ADDR
      EditBox 55, 105, 135, 15, new_ADDR_line_1
      EditBox 55, 125, 135, 15, new_ADDR_line_2
      EditBox 55, 145, 135, 15, new_addr_city
      EditBox 55, 165, 35, 15, new_addr_zip
      EditBox 165, 165, 25, 15, new_addr_state
	  DropListBox 55, 185, 35, 15, "Select One:"+chr(9)+"Aitkin"+chr(9)+"Anoka"+chr(9)+"Becker"+chr(9)+"Beltrami"+chr(9)+"Benton"+chr(9)+"Big Stone"+chr(9)+"Blue Earth"+chr(9)+"Brown"+chr(9)+"Carlton"+chr(9)+"Carver"+chr(9)+"Cass"+chr(9)+"Chippewa"+chr(9)+"Chisago"+chr(9)+"Clay"+chr(9)+"Clearwater"+chr(9)+"Cook"+chr(9)+"Cottonwood"+chr(9)+"Crow Wing"+chr(9)+"Dakota"+chr(9)+"Dodge"+chr(9)+"Douglas"+chr(9)+"Faribault"+chr(9)+"Fillmore"+chr(9)+"Freeborn"+chr(9)+"Goodhue"+chr(9)+"Grant"+chr(9)+"Hennepin"+chr(9)+"Houston"+chr(9)+"Hubbard"+chr(9)+"Isanti"+chr(9)+"Itasca"+chr(9)+"Jackson"+chr(9)+"Kanabec"+chr(9)+"Kandiyohi"+chr(9)+"Kittson"+chr(9)+"Koochiching"+chr(9)+"Lac Qui Parle"+chr(9)+"Lake"+chr(9)+"Lake Of Woods"+chr(9)+"Le Sueur"+chr(9)+"Lincoln"+chr(9)+"Lyon"+chr(9)+"Mcleod"+chr(9)+"Mahnomen"+chr(9)+"Marshall"+chr(9)+"Martin"+chr(9)+"Meeker"+chr(9)+"Mille Lacs"+chr(9)+"Morrison"+chr(9)+"Mower"+chr(9)+"Murray"+chr(9)+"Nicollet"+chr(9)+"Nobles"+chr(9)+"Norman"+chr(9)+"Olmsted"+chr(9)+"Otter Tail"+chr(9)+"Pennington"+chr(9)+"Pine"+chr(9)+"Pipestone"+chr(9)+"Polk"+chr(9)+"Pope"+chr(9)+"Ramsey"+chr(9)+"Red Lake"+chr(9)+"Redwood"+chr(9)+"Renville"+chr(9)+"Rice"+chr(9)+"Rock"+chr(9)+"Roseau"+chr(9)+"St. Louis"+chr(9)+"Scott"+chr(9)+"Sherburne"+chr(9)+"Sibley"+chr(9)+"Stearns"+chr(9)+"Steele"+chr(9)+"Stevens"+chr(9)+"Swift"+chr(9)+"Todd"+chr(9)+"Traverse"+chr(9)+"Wabasha"+chr(9)+"Wadena"+chr(9)+"Waseca"+chr(9)+"Washington"+chr(9)+"Watonwan"+chr(9)+"Wilkin"+chr(9)+"Winona"+chr(9)+"Wright"+chr(9)+"Yellow Medicine"+chr(9)+"Out-of-State", county_code
	  DropListBox 155, 185, 35, 15, "Select One:"+chr(9)+"Residence"+chr(9)+"Mailing"+chr(9)+"Both"+chr(9)+"Unknown", residence_addr
	  DropListBox 55, 200, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", homeless_addr
	  DropListBox 155, 200, 35, 15, "Select One:"+chr(9)+"Own Housing: Lease, Mortgage or Roommate"+chr(9)+"Family/Friends Due to Economic Hardship"+chr(9)+"Service Provider-Foster Care Group Home"+chr(9)+"Hospital/Treatment/Detox/Nursing Home"+chr(9)+"Jail/Prison/Juvenile Detention Center"+chr(9)+"Hotel/Motel"+chr(9)+"Emergency Shelter"+chr(9)+"Place Not Meant for housing"+chr(9)+"Declined"+chr(9)+"Unknown", living_situation
	  DropListBox 125, 215, 65, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", reservation_name
	  DropListBox 55, 215, 35, 15, "Select One:"+chr(9)+"Bois Forte-Deer Creek"+chr(9)+"Bois Forte-Nett Lake"+chr(9)+"Bois Forte-Vermillion Lk"+chr(9)+"Fond du Lac"+chr(9)+"Grand Portage"+chr(9)+"Leach Lake"+chr(9)+"Lower Sioux"+chr(9)+"Mille Lacs"+chr(9)+"Prairie Island Community"+chr(9)+"Red Lake"+chr(9)+"Shakopee Mdewakanton"+chr(9)+"Upper Sioux"+chr(9)+"White Earth", reservation_addr
	  DropListBox 125, 240, 65, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", METS_ADDR
	  EditBox 135, 255, 55, 15, MNsure_number
	  EditBox 50, 275, 140, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 75, 275, 50, 15
	    CancelButton 140, 275, 50, 15
	  Text 35, 150, 20, 10, "City:"
	  Text 140, 170, 20, 10, "State:"
	  Text 50, 95, 85, 10, "Script to update address:"
	  Text 10, 190, 30, 10, "County:"
	  Text 20, 170, 35, 10, "Zip code:"
	  GroupBox 5, 80, 190, 155, "New Address:"
	  Text 30, 110, 20, 10, "Street:"
	  Text 15, 130, 35, 10, "Apt/Room:"
	  Text 5, 275, 40, 10, "Other notes:"
	  Text 100, 220, 25, 10, "Name:"
	  GroupBox 5, 5, 190, 50, "Address in MAXIS:"
	  Text 10, 220, 45, 10, "Reservation:"
	  Text 5, 240, 95, 10, "METS correspondence sent:"
	  Text 10, 205, 35, 10, "Homeless:"
	  GroupBox 5, 55, 190, 25, "Verification Request Form(s) Sent:"
	  Text 100, 190, 40, 10, "Is address:"
	  Text 100, 205, 55, 10, "Living situation:"
	  Text 5, 260, 70, 10, "METS case number:"
	EndDialog

    DO
    	DO
    		err_msg = ""
    		DIALOG returned_mail_update_addr
    			IF ButtonPressed = 0 THEN stopscript
    			IF MNsure_ADDR = "YES" THEN
    				IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
    			END IF
    			IF MNsure_ADDR = "Select One:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    		LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    CALL check_for_MAXIS(False)
END IF

IF homeless_addr = "YES" THEN homeless_addr_code = "Y"
IF homeless_addr = "NO" THEN homeless_addr_code = "N"
IF reservation_name = "YES" THEN reservation_name = "Y"
IF reservation_name = "NO" THEN reservation_name = "N"

IF county_code = "Aitkin" THEN county_code_number = "01"
IF county_code = "Anoka" THEN county_code_number = "02"
IF county_code = "Becker" THEN county_code_number = "03"
IF county_code = "Beltrami" THEN county_code_number = "04"
IF county_code = "Benton" THEN county_code_number = "05"
IF county_code = "BigStone" THEN county_code_number = "06"
IF county_code = "BlueEarth" THEN county_code_number = "07"
IF county_code = "Brown" THEN county_code_number = "08"
IF county_code = "Carlton" THEN county_code_number = "09"
IF county_code = "Carver" THEN county_code_number = "10"
IF county_code = "Cass" THEN county_code_number = "11"
IF county_code = "Chippewa" THEN county_code_number = "12"
IF county_code = "Chisago" THEN county_code_number = "13"
IF county_code = "Clay" THEN county_code_number = "14"
IF county_code = "Clearwater" THEN county_code_number = "15"
IF county_code = "Cook" THEN county_code_number = "16"
IF county_code = "Cottonwood" THEN county_code_number = "17"
IF county_code = "CrowWing" THEN county_code_number = "18"
IF county_code = "Dakota" THEN county_code_number = "19"
IF county_code = "Dodge" THEN county_code_number = "20"
IF county_code = "Douglas" THEN county_code_number = "21"
IF county_code = "Faribault" THEN county_code_number = "22"
IF county_code = "Fillmore" THEN county_code_number = "23"
IF county_code = "Freeborn" THEN county_code_number = "24"
IF county_code = "Goodhue" THEN county_code_number = "25"
IF county_code = "Grant" THEN county_code_number = "26"
IF county_code = "Hennepin" THEN county_code_number = "27"
IF county_code = "Houston" THEN county_code_number = "28"
IF county_code = "Hubbard" THEN county_code_number = "29"
IF county_code = "Isanti" THEN county_code_number = "30"
IF county_code = "Itasca" THEN county_code_number = "31"
IF county_code = "Jackson" THEN county_code_number = "32"
IF county_code = "Kanabec" THEN county_code_number = "33"
IF county_code = "Kandiyohi" THEN county_code_number = "34"
IF county_code = "Kittson" THEN county_code_number = "35"
IF county_code = "Koochiching" THEN county_code_number = "36"
IF county_code = "Lac QuiPar" THEN county_code_number = "37"
IF county_code = "Lake" THEN county_code_number = "38"
IF county_code = "Lake Of Woods" THEN county_code_number = "39"
IF county_code = "LeSueur" THEN county_code_number = "40"
IF county_code = "Lincoln" THEN county_code_number = "41"
IF county_code = "Lyon" THEN county_code_number = "42"
IF county_code = "Mcleod" THEN county_code_number = "43"
IF county_code = "Mahnomen" THEN county_code_number = "44"
IF county_code = "Marshall" THEN county_code_number = "45"
IF county_code = "Martin" THEN county_code_number = "46"
IF county_code = "Meeker" THEN county_code_number = "47"
IF county_code = "MilleLacs" THEN county_code_number = "48"
IF county_code = "Morrison" THEN county_code_number = "49"
IF county_code = "Mower" THEN county_code_number = "50"
IF county_code = "Murray" THEN county_code_number = "51"
IF county_code = "Nicollet" THEN county_code_number = "52"
IF county_code = "Nobles" THEN county_code_number = "53"
IF county_code = "Norman" THEN county_code_number = "54"
IF county_code = "Olmsted" THEN county_code_number = "55"
IF county_code = "OtterTail" THEN county_code_number = "56"
IF county_code = "Pennington" THEN county_code_number = "57"
IF county_code = "Pine" THEN county_code_number = "58"
IF county_code = "Pipestone" THEN county_code_number = "59"
IF county_code = "Polk" THEN county_code_number = "60"
IF county_code = "Pope" THEN county_code_number = "61"
IF county_code = "Ramsey" THEN county_code_number = "62"
IF county_code = "RedLake" THEN county_code_number = "63"
IF county_code = "Redwood" THEN county_code_number = "64"
IF county_code = "Renville" THEN county_code_number = "65"
IF county_code = "Rice" THEN county_code_number = "66"
IF county_code = "Rock" THEN county_code_number = "67"
IF county_code = "Roseau" THEN county_code_number = "68"
IF county_code = " St.Louis" THEN county_code_number = "69"
IF county_code = "Scott" THEN county_code_number = "70"
IF county_code = "Sherburne" THEN county_code_number = "71"
IF county_code = "Sibley" THEN county_code_number = "72"
IF county_code = "Stearns" THEN county_code_number = "73"
IF county_code = "Steele" THEN county_code_number = "74"
IF county_code = "Stevens" THEN county_code_number = "75"
IF county_code = "Swift" THEN county_code_number = "76"
IF county_code = "Todd" THEN county_code_number = "77"
IF county_code = "Traverse" THEN county_code_number = "78"
IF county_code = "Wabasha" THEN county_code_number = "79"
IF county_code = "Wadena" THEN county_code_number = "80"
IF county_code = "Waseca" THEN county_code_number = "81"
IF county_code = "Washington" THEN county_code_number = "82"
IF county_code = "Watonwan" THEN county_code_number = "83"
IF county_code = "Wilkin" THEN county_code_number = "84"
IF county_code = "Winona" THEN county_code_number = "85"
IF county_code = "Wright" THEN county_code_number = "86"
IF county_code = "YellowMedine" THEN county_code_number = "87"
IF county_code = "Out-ofState" THEN county_code_number = "89"

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

IF reservation_addr = "Bois Forte-Deer Creek" THEN rez_code = "BD"
IF reservation_addr = "Bois Forte-Nett Lake" THEN rez_code = "BN"
IF reservation_addr = "Bois Forte-Vermillion Lk" THEN rez_code = "BV"
IF reservation_addr = "Fond du Lac" THEN rez_code = "FL"
IF reservation_addr = "Grand Portage" THEN rez_code = "GP"
IF reservation_addr = "Leach Lake" THEN rez_code = "LL"
IF reservation_addr = "Lower Sioux" THEN rez_code = "LS"
IF reservation_addr = "Mille Lacs" THEN rez_code = "ML"
IF reservation_addr = "Prairie Island Community" THEN rez_code = "PL"
IF reservation_addr = "Red Lake" THEN rez_code = "RL"
IF reservation_addr = "Shakopee Mdewakanton" THEN rez_code = "SM"
IF reservation_addr = "Upper Sioux" THEN rez_code = "US"
IF reservation_addr = "White Earth" THEN rez_code = "WE"


IF update_ADDR = "YES" THEN
	PF9
	IF residence_addr = "Residence" THEN
	    EMwritescreen new_addr_line_one, 20, 6, 43
	    EMwritescreen ADDR_line_two, 20, 7, 43
	    EMwritescreen new_addr_city, 15, 8, 43
	    EMwritescreen new_addr_state, 8, 66		'Defaults to MN for all cases at this time
	    EMwritescreen new_addr_zip, 5, 9, 43
	END IF
	EMwritescreen county_code_number, 9, 66
	EMwritescreen "OT", 9, 74
	EMwritescreen homeless_addr_code, 10, 43

	EMwritescreen reservation_name, 10, 74
	EMwritescreen rez_code, 11, 74
	EMwritescreen living_situation_code, 11, 43

	EMwritescreen new_addr_line_one, 13, 43
	EMwritescreen ADDR_mailing_addr_line_two, 14, 43
	EMwritescreen new_addr_city, 15, 43
	EMwritescreen new_addr_state, 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
	EMwritescreen new_addr_zip, 5, 16, 52
END IF

IF ADDR_actions = "Client has not responded to request for verif" THEN

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG

'Setting some variables for the loop
CASH_STATUS = FALSE 'overall variable'
CCA_STATUS = FALSE
DW_STATUS = FALSE 'Diversionary Work Program'
ER_STATUS = FALSE
SNAP_STATUS = FALSE
GA_STATUS = FALSE 'General Assistance'
GRH_STATUS = FALSE
HC_STATUS = FALSE
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67
EMReadScreen cca_prog_check, 2, 14, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND"  THEN SNAP_STATUS = TRUE
IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
IF hc_status_check = "ACTV" or hc_status_check = "PEND"  THEN HC_STATUS = TRUE
IF cca_status_check = "ACTV" or cca_status_check = "PEND"  THEN CCA_STATUS = TRUE
'Logic to determine if Cash is active
If cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "RC" or cash1_prog_check = "MS" THEN
	IF cash1_status_check = "ACTV" THEN CASH_STATUS = TRUE
	IF cash1_status_check = "PEND" THEN CASH_STATUS = TRUE
	IF cash1_status_check = "INAC" THEN CASH_STATUS = FALSE
	IF cash1_status_check = "SUSP" THEN CASH_STATUS = FALSE
	IF cash1_status_check = "DENY" THEN CASH_STATUS = FALSE
	IF cash1_status_check = ""     THEN CASH_STATUS = FALSE
END IF
If cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "RC" or cash2_prog_check = "MS" THEN
	IF cash2_status_check = "ACTV" THEN CASH_STATUS = TRUE
	IF cash2_status_check = "PEND" THEN CASH_STATUS = TRUE
	IF cash2_status_check = "INAC" THEN CASH_STATUS = FALSE
	IF cash2_status_check = "SUSP" THEN CASH_STATUS = FALSE
	IF cash2_status_check = "DENY" THEN CASH_STATUS = FALSE
	IF cash2_status_check = ""     THEN CASH_STATUS = FALSE
END IF
'per POLI/TEM this only pretains to cash and snap '
IF CASH_STATUS = TRUE or SNAP_STATUS = TRUE THEN
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
			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYesNo + vbQuestion, "Message handling")
			IF maxis_error_check = vbNO THEN
				case_note_only = FALSE 'this will case note only'
			END IF
			IF maxis_error_check= vbYES THEN
				case_note_only = TRUE 'this will update the panels and case note'
			END IF
		END IF
    ELSE
		IF case_note_only = FALSE THEN


    			EmReadScreen open_cash1, 2, 6, 43
				EmReadScreen open_cash2, 2, 8, 43
				EmReadScreen open_grh, 2, 10, 43
				EmReadScreen open_snap, 2, 12, 43
				IF open_cash1 <> "" THEN EMWriteScreen "3", 6, 58
				IF open_cash2 <> "" THEN EMWriteScreen "3", 8, 58
				IF open_grh <> "" THEN EMWriteScreen "3", 10, 58
				IF open_snap <> "" THEN EMWriteScreen "3", 12, 58
				TRANSMIT
   		END IF
	END IF
'assigns a value to the ADDR_status variable based on the value of the complete variable
IF forwarding_ADDR = "Yes" THEN ADDR_status = "a forwarding ADDR."
IF forwarding_ADDR = "No" THEN ADDR_status = "no forwarding ADDR."

'assigns a value to the MNsure variable based on the value of MNsure_active
IF MNsure_active = "Yes" THEN MNsure = "MNsure case"
IF MNsure_active = "No" THEN MNsure = "Non-MNsure"


'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
call check_for_MAXIS (false)

'starts a blank case note
call start_a_blank_case_note

case_note_header = "Mail has been returned with forwarding address"

  CheckBox 10, 65, 50, 10, "DHS-2919A", verifA_sent_checkbox
  CheckBox 70, 65, 45, 10, "DHS-2952", SHEL_form_sent_checkbox
  CheckBox 125, 65, 45, 10, "DHS-2402", CRF_sent_checkbox


"New Address:",

call write_variable_in_CASE_NOTE("***Returned Mail received on: " & date_received & ".")
call write_bullet_and_variable_in_CASE_NOTE("Street:", new_ADDR_line_1)
call write_bullet_and_variable_in_CASE_NOTE("Apt/Room:", new_ADDR_line_2)
call write_bullet_and_variable_in_CASE_NOTE("City:", new_addr_city)
call write_bullet_and_variable_in_CASE_NOTE("State:", new_addr_state)
call write_bullet_and_variable_in_CASE_NOTE("County:", county_code)
call write_bullet_and_variable_in_CASE_NOTE("Zip code:", new_addr_zip)
call write_bullet_and_variable_in_CASE_NOTE()
call write_bullet_and_variable_in_CASE_NOTE("Previous address in MAXIS:", maxis_addr)
call write_bullet_and_variable_in_CASE_NOTE("Reservation:", reservation_addr)
call write_bullet_and_variable_in_CASE_NOTE("METS correspondence sent:", METS_ADDR)
call write_bullet_and_variable_in_CASE_NOTE("Homeless:", homeless_addr)
call write_bullet_and_variable_in_CASE_NOTE("Verification Request Form(s) Sent:",)
call write_bullet_and_variable_in_CASE_NOTE("Is address:", reservation_name)
call write_bullet_and_variable_in_CASE_NOTE("Living situation:", living_situation)
call write_variable_in_CASE_NOTE ("---")
call write_bullet_and_variable_in_CASE_NOTE("METS case number:", MNsure_number)
If verifA_sent_checkbox = CHECKED then call write_variable_in_CASE_NOTE("* Verification Request Form sent.")
If SHEL_form_sent_checkbox = CHECKED then call write_variable_in_CASE_NOTE("* Shelter Verification Form sent.")
If CRF_sent_checkbox = CHECKED then call write_variable_in_CASE_NOTE("* Change Report Form sent.")






call write_variable_in_CASE_NOTE("* Address updated to: " & new_addr_line_one)
THEN call write_variable_in_CASE_NOTE("                      " & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
THEN call write_variable_in_CASE_NOTE("* New ADDR is in " & new_COUNTY & " COUNTY.")

call write_bullet_and_variable_in_CASE_NOTE("Returned Mail resent", mail_resent)
If returned_mail_resent_list = "Yes" then call write_variable_in_CASE_NOTE("* Returned Mail resent to client.")
If returned_mail_resent_list = "No" then call write_variable_in_CASE_NOTE("* Returned Mail not resent to client.")



call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Checks if this is a MNsure case and pops up a message box with instructions if the ADDR is incorrect.
IF MNsure_active = "Yes" and MNsure_ADDR = "No" THEN MsgBox "Please update the MNsure ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. HPU Case Manitenance - Action Needed Log)"


'Checks if a DHS2919A mailed and sets a TIKL for the return of the info.
IF verifA_sent_checkbox = CHECKED THEN
	call navigate_to_MAXIS_screen("dail", "writ")

	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)

	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("ADDR verification requested via 2919A after returned mail being rec'd should have returned by now. If not received, take appropriate action." )
	transmit
	PF3

	'Success message
	MsgBox "Success! TIKL has been sent for 10 days from now for the ADDR verification requested via 2919A."

End if

script_end_procedure("")
