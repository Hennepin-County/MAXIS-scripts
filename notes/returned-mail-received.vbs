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
call changelog_update("06/06/2019", "Initial version. Rewitten per POLI/TEMP.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'------------------------------------------------------------------------------------------------------------------------------THE SCRIPT
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
BeginDialog case_number_dlg, 0, 0, 211, 85, "Returned Mail"
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
    	DIALOG case_number_dlg
    		IF ButtonPressed = 0 THEN stopscript
    		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
			If isdate(date_received) = FALSE and Cdate(date_received) > cdate(date) = TRUE THEN  err_msg = err_msg & vbnewline & "* You must enter an actual date that is not in the future and is in the footer month that you are working in."
			'If isdate(date_received) = FALSE then err_msg = err_msg & vbnewline & "* Please enter a date (--/--/--) in the footer month that you are working in."
			IF ADDR_actions = "Select:" THEN err_msg = err_msg & vbCr & "Please chose an action for the returned mail."
    		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
CALL check_for_MAXIS(False)


'MAXIS_footer_month = right("00" & DatePart("m", date_received), 2)
MAXIS_footer_month = DatePart("m", date_received)
MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)

'MAXIS_footer_year = right(DatePart("yyyy", date_received), 2)
MAXIS_footer_year = DatePart("yyyy", date_received)
MAXIS_footer_year = right("yyyy" & MAXIS_footer_year, 2)

msgbox MAXIS_footer_year


CALL navigate_to_MAXIS_screen("STAT", "SELF")		'Goes to STAT/PROG
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" THEN
	EmWriteScreen MAXIS_footer_month, 04, 43
	EmWriteScreen MAXIS_footer_year, 04, 49
	TRANSMIT
END IF

CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
EMReadScreen err_msg, 7, 24, 02
IF err_msg = "BENEFIT" THEN	script_end_procedure_with_error_report ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

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
EMreadscreen resi_addr_line_one, 20, 6, 43
resi_addr_line_one = replace(resi_addr_line_one, "_", "")
EMreadscreen resi_addr_line_two, 20, 7, 43
resi_addr_line_two = replace(resi_addr_line_two, "_", "")
EMreadscreen resi_addr_city, 15, 8, 43
resi_addr_city = replace(resi_addr_city, "_", "")
EMreadscreen resi_addr_state, 2, 8, 66
EMreadscreen resi_addr_zip, 5, 9, 43
resi_addr_zip = replace(resi_addr_zip, "_", "")
EMreadscreen addr_county, 2, 9, 66
EMreadscreen addr_verif, 2, 9, 74
EMreadscreen addr_homeless, 1, 10, 43
'EMreadscreen ADDR_reservation, 1, 10, 74
EMreadscreen mailing_addr_line_one, 20, 13, 43
mailing_addr_line_one = replace(mailing_addr_line_one, "_", "")
EMreadscreen mailing_addr_line_two, 20, 14, 43
mailing_addr_line_two = replace(mailing_addr_line_two, "_", "")
EMreadscreen mailing_addr_city, 15, 15, 43
mailing_addr_city = replace(mailing_addr_city, "_", "")
EmReadScreen mailing_addr_state, 2, 16, 43
EMreadscreen mailing_addr_zip, 5, 16, 52
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

IF ADDR_actions = "No forwarding address provided" THEN
    BeginDialog no_forward_addr, 0, 0, 186, 215, "No forwarding address provided"
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
	    	DIALOG no_forward_addr
	    		IF ButtonPressed = 0 THEN stopscript
				IF verifA_sent_checkbox = UNCHECKED and CRF_sent_checkbox = UNCHECKED and SHEL_form_sent_checkbox= UNCHECKED THEN err_msg = err_msg & vbCr & "Please select the verifcation requested and ensure forms are sent in ECF."
	    		IF mets_addr_correspondence = "YES" THEN
					IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
				END IF
				IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
	    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    	LOOP UNTIL err_msg = ""
	    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	CALL check_for_MAXIS(False)
END IF

IF ADDR_actions = "Forwarding address outside MN" THEN county_code = "Out of State"
IF ADDR_actions = "Forwarding address in MN"  then new_addr_state = "MN"

IF ADDR_actions = "Forwarding address in MN" or ADDR_actions = "Forwarding address outside MN" THEN
    BeginDialog returned_mail_update_addr, 0, 0, 206, 305, "Mail has been returned with forwarding address"
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
      DropListBox 130, 175, 65, 15, "Select:"+chr(9)+"Aitkin"+chr(9)+"Anoka"+chr(9)+"Becker"+chr(9)+"Beltrami"+chr(9)+"Benton"+chr(9)+"Big Stone"+chr(9)+"Blue Earth"+chr(9)+"Brown"+chr(9)+"Carlton"+chr(9)+"Carver"+chr(9)+"Cass"+chr(9)+"Chippewa"+chr(9)+"Chisago"+chr(9)+"Clay"+chr(9)+"Clearwater"+chr(9)+"Cook"+chr(9)+"Cottonwood"+chr(9)+"Crow     Wing"+chr(9)+"Dakota"+chr(9)+"Dodge"+chr(9)+"Douglas"+chr(9)+"Faribault"+chr(9)+"Fillmore"+chr(9)+"Freeborn"+chr(9)+"Goodhue"+chr(9)+"Grant"+chr(9)+"Hennepin"+chr(9)+"Houston"+chr(9)+"Hubbard"+chr(9)+"Isanti"+chr(9)+"Itasca"+chr(9)+"Jackson"+chr(9)+"Kanabec"+chr(9)+"Kandiyohi"+chr(9)+"Kittson"+chr(9)+"Koochiching"+chr(9)+"Lac Qui Parle"+chr(9)+"Lake"+chr(9)+"Lake Of Woods"+chr(9)+"Le     Sueur"+chr(9)+"Lincoln"+chr(9)+"Lyon"+chr(9)+"Mcleod"+chr(9)+"Mahnomen"+chr(9)+"Marshall"+chr(9)+"Martin"+chr(9)+"Meeker"+chr(9)+"Mille Lacs"+chr(9)+"Morrison"+chr(9)+"Mower"+chr(9)+"Murray"+chr(9)+"Nicollet"+chr(9)+"Nobles"+chr(9)+"Norman"+chr(9)+"Olmsted"+chr(9)+"Otter Tail"+chr(9)+"Pennington"+chr(9)+"Pine"+chr(9)+"Pipestone"+chr(9)+"Polk"+chr(9)+"Pope"+chr(9)+"Ramsey"+chr(9)+"Red Lake"+chr(9)+"Redwood"+chr(9)+"Renville"+chr(9)+"Rice"+chr(9)+"Rock"+chr(9)+"Roseau"+chr(9)+"St.     Louis"+chr(9)+"Scott"+chr(9)+"Sherburne"+chr(9)+"Sibley"+chr(9)+"Stearns"+chr(9)+"Steele"+chr(9)+"Stevens"+chr(9)+"Swift"+chr(9)+"Todd"+chr(9)+"Traverse"+chr(9)+"Wabasha"+chr(9)+"Wadena"+chr(9)+"Waseca"+chr(9)+"Washington"+chr(9)+"Watonwan"+chr(9)+"Wilkin"+chr(9)+"Winona"+chr(9)+"Wright"+chr(9)+"Yellow Medicine"+chr(9)+"Out of State", county_code
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
    		DIALOG returned_mail_update_addr
    			IF ButtonPressed = 0 THEN stopscript
				IF new_addr_line_one = "" THEN err_msg = err_msg & vbCr & "Please complete the street address the client in now living at."
				IF new_addr_city = "" THEN err_msg = err_msg & vbCr & "Please complete the city in which the client in now living."
				IF new_addr_state = "" THEN err_msg = err_msg & vbCr & "Please complete the state in which the client in now living."
				IF new_addr_zip = "" THEN err_msg = err_msg & vbCr & "Please complete the city in which the client in now living."
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
    CALL check_for_MAXIS(False)
END IF

IF homeless_addr = "YES" THEN homeless_addr_code = "Y"
IF homeless_addr = "NO" THEN homeless_addr_code = "N"
IF reservation_addr = "YES" THEN reservation_addr_code = "Y"
IF reservation_addr = "NO" THEN reservation_addr_code = "N"

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
IF county_code = "Out of State" THEN county_code_number = "89"

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
	'MsgBox MAXIS_footer_month & " " & MAXIS_footer_year
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
	'MsgBox " Residence -how did we do"
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
	'MsgBox " Residence -how did we do"
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
	TRANSMIT
	EMReadScreen error_msg, 75, 24, 2
	error_msg = TRIM(error_msg)
	IF error_msg <> "" THEN
		MsgBox "*** NOTICE!!! ***" & vbNewLine & error_msg & vbNewLine
		IF error_msg = "WARNING: EFFECTIVE DATE HAS CHANGED - REVIEW LIVING SITUATION" THEN
			TRANSMIT
		END IF
	END IF
END IF

IF ADDR_actions = "Client has not responded to request" THEN
    BeginDialog date_rcvd_dialog, 0, 0, 151, 100, "Client has not responded to request"
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
    		DIALOG date_rcvd_dialog
    			IF ButtonPressed = 0 THEN stopscript
    			If isdate(date_requested) = FALSE and Cdate(date_requested) > cdate(date) = TRUE THEN  err_msg = err_msg & vbnewline & "* Please enter the date verifcations were requested."
    			IF ECF_review_checkbox <> CHECKED THEN  err_msg = err_msg & vbnewline & "* Please review ECF to ensure the requested verifications are not on file."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    		LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    CALL check_for_MAXIS(False)

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
			'MsgBox "did we edit PACT"
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
'msgbox"what did we do?"

pending_verifs = ""
IF verifA_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Verification Request, "
IF SHEL_form_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "SVF, "
IF CRF_sent_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Change Request Form, "
IF return_mail_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Returned Mail/Other, "
'-------------------------------------------------------------------trims excess spaces of pending_verifs
pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(False)

'checking to make sure case is out of background & gets to STAT/BUDG
Call MAXIS_background_check

'starts a blank case note
call start_a_blank_case_note
call write_variable_in_CASE_NOTE("Returned mail received-" & ADDR_actions & " for " & active_programs & "")
IF ADDR_actions <> "Client has not responded to request" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Received on", date_received)
IF ADDR_actions = "Forwarding address in MN" or ADDR_actions = "Forwarding address outside MN" THEN
    CALL write_bullet_and_variable_in_CASE_NOTE("Living situation", living_situation)
    CALL write_bullet_and_variable_in_CASE_NOTE("Homeless", homeless_addr)
	CALL write_bullet_and_variable_in_case_note("Verification(s) received", pending_verifs)
    CALL write_variable_in_CASE_NOTE("* Mailing address updated:  " & new_addr_line_one)
    CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
    CALL write_variable_in_CASE_NOTE("                            " & county_code & " COUNTY.")
	IF reservation_name <> "Select:" or reservation_name <> "N/A" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
    'CALL write_variable_in_CASE_NOTE("---")
ELSE
	CALL write_bullet_and_variable_in_case_note("Verification(s) requested", pending_verifs)
END IF
CALL write_variable_in_CASE_NOTE ("---")
'IF mailing_addr_line_one <> "" THEN CALL write_variable_in_CASE_NOTE("* No mailing address entered in Maxis")
IF mailing_addr_line_one <> "" THEN
	CALL write_variable_in_CASE_NOTE("* Previous mailing address: " & mailing_addr_line_one)
	CALL write_variable_in_CASE_NOTE("                           " & mailing_addr_line_two & " " & mailing_addr_city & " " & mailing_addr_state & " " & mailing_addr_zip)
ELSE
	CALL write_variable_in_CASE_NOTE("* Previous residential address: " & resi_addr_line_one)
	CALL write_variable_in_CASE_NOTE("                               " & resi_addr_line_two & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip)
END IF
IF mets_addr_correspondence <> "Select:" CALL write_bullet_and_variable_in_CASE_NOTE("METS correspondence sent", mets_addr_correspondence)
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

'Checks if a DHS2919A mailed and sets a TIKL for the return of the info.
IF ADDR_actions <> "Client has not responded to request" THEN
	call navigate_to_MAXIS_screen("dail", "writ")
	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("Returned mail rec'd contact from the client should have occured regarding address change. If no response-verbal or written, please take appropriate action." )
	TRANSMIT
	PF3
End if

IF ADDR_actions <> "Client has not responded to request" THEN script_end_procedure_with_error_report("Success! TIKL has been set for 10 days for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over the telephone or in person, the unit is not required to also report the change on a Change Report from. ")
IF ADDR_actions = "Client has not responded to request" THEN script_end_procedure_with_error_report("Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter a worker comment in SPEC/WCOM.")
