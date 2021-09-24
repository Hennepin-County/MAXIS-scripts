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
call changelog_update("03/12/2021", "GitHub Issue #309 Updated handling for current address confirmation.", "MiKayla Handley")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("02/13/2020", "Updated the zip code to only allow for 5 characters.", "MiKayla Handley, Hennepin County")
call changelog_update("06/06/2019", "Initial version. Re-written per POLI/TEMP.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

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

' CALL read_ADDR_panel(addr_eff_date, addr_future_date, resi_addr_line_one, resi_addr_line_two,              resi_addr_city, resi_addr_state, resi_addr_zip,                                                                                              mail_line_one, mail_line_two,                   mail_city_line, mail_state_line, mail_zip_line, living_situation, living_sit_line, homeless_line, addr_phone_1A)
Call access_ADDR_panel("READ", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_county, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to navigate'
IF trim(case_invalid_error) <> "" THEN script_end_procedure("*** NOTICE!!! ***" & vbCr & case_invalid_error & vbCr & "Please resolve for the script to continue.")
'-------------------------------------------------------------------------------------------------DIALOG
mailing_address_confirmed = "NO"
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 566, 105, "Returned Mail Information"
  DropListBox 220, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", residential_address_confirmed
  DropListBox 500, 15, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", mailing_address_confirmed
  EditBox 75, 85, 205, 15, notes_on_address
  ButtonGroup ButtonPressed
    PushButton 285, 85, 50, 15, "CASE/ADHI", ADHI_button
    PushButton 355, 85, 70, 15, "HSR Manual/SNAP", HSR_manual_button
    OkButton 445, 85, 55, 15
    CancelButton 505, 85, 55, 15
  Text 10, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
  ' Text 15, 30, 200, 10, resi_addr_line_one
  Text 15, 40, 200, 10, resi_addr_street_full
  Text 15, 50, 205, 10, resi_addr_city &  " , "  & resi_addr_state & " , "   & resi_addr_zip
  Text 10, 90, 65, 10, "Notes on Address:"
  GroupBox 285, 5, 275, 75, "Mailing Address"
  Text 290, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
  ' Text 295, 30, 200, 10, mail_line_one
  Text 295, 40, 200, 10, mail_street_full
  If mail_city_line <> "" Then Text 295, 50, 205, 10, mail_city_line & " , "  & mail_state_line &  " , "  & mail_zip_line
  GroupBox 5, 5, 275, 75, "Residential Address in Maxis"
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
		IF residential_address_confirmed = "Select One:" THEN err_msg = err_msg & vbCr & "Please confirm if this the address that the agency attempted to deliver mail to."
		IF mail_line_one <> "" THEN
			IF mailing_address_confirmed = "Select One:" THEN err_msg = err_msg & vbCr & "Please confirm if the mailing address is the address that the agency attempted to deliver mail to."
		END IF

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
		reservation_addr = "No"
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
          DropListBox 40, 60, 155, 15, "Select:"+chr(9)+"01 - Own home, lease or roomate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown", living_situation
          EditBox 40, 80, 155, 15, new_addr_line_one
          EditBox 40, 100, 155, 15, new_addr_line_two
          EditBox 40, 120, 155, 15, new_addr_city
          EditBox 40, 140, 20, 15, new_addr_state
          EditBox 155, 140, 40, 15, new_addr_zip
          DropListBox 55, 165, 40, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", homeless_addr
          DropListBox 55, 185, 40, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", reservation_addr
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
    			IF homeless_addr = "Select:" THEN err_msg = err_msg & vbCr & "Please advise whether the client has reported homelessness."
    			IF living_situation = "Select:" THEN err_msg = err_msg & vbCr & "Please select the client's living situation - Unknown should not if it is avoidable."
    			IF reservation_addr = "Select:" THEN err_msg = err_msg & vbCr & "Please select if client is living on the reservation."
    			IF reservation_addr = "Yes" THEN
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

    	new_addr_line_one = trim(new_addr_line_one)
        new_addr_line_two = trim(new_addr_line_two)
        new_addr_city = trim(new_addr_city)
        new_addr_state = trim(new_addr_state)
        new_addr_zip = trim(new_addr_zip)

		living_situation_code = left(living_situation, 2)

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
		IF residential_address_confirmed = "YES" THEN
			original_resi_addr_line_one = resi_addr_line_one
			original_resi_addr_line_two = resi_addr_line_two
			original_resi_addr_city = resi_addr_city
			original_resi_addr_state = resi_addr_state
			original_resi_addr_zip = resi_addr_zip

			resi_addr_line_one = new_addr_line_one
			resi_addr_line_two = new_addr_line_two
			resi_addr_city = new_addr_city
			resi_addr_state = new_addr_state
			resi_addr_zip = new_addr_zip
		End If
		IF mailing_address_confirmed = "YES" THEN
			original_mail_line_one = mail_line_one
			original_mail_line_two = mail_line_two
			original_mail_city_line = mail_city_line
			original_mail_state_line = mail_state_line
			original_mail_zip_line = mail_zip_line

			mail_line_one = new_addr_line_one
			mail_line_two = new_addr_line_two
			mail_city_line = new_addr_city
			mail_state_line = new_addr_state
			mail_zip_line = new_addr_zip
		End If

		IF residential_address_confirmed = "YES" OR mailing_address_confirmed = "YES" THEN
			begining_of_footer_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
			begining_of_footer_month = DateAdd("d", 0, begining_of_footer_month)
			If DateDiff("d", begining_of_footer_month, addr_eff_date) > 0 Then begining_of_footer_month = addr_eff_date
			Call access_ADDR_panel("WRITE", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, county_code, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
		End If
	END IF

    IF ADDR_actions = "no response received" THEN
		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
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
		'IF case_pending = TRUE THEN 'TODO if both cash one and cash two are active the pact panel will need to be updated manually

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

        'per POLI/TEMP this only pertains to active cash and snap '
        IF case_active = TRUE THEN
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
        		'IF cash2_active = TRUE THEN EMWriteScreen "3", 8, 58 'TODO need to explain this is the reason for reading PROG'
        		IF SNAP_active  = TRUE THEN EMWriteScreen "4", 12, 58
        		IF grh_active = TRUE THEN EMWriteScreen "3", 10, 58
        		TRANSMIT
        		EMReadScreen pop_upmsg, 7, 11, 08
        		IF pop_upmsg = "WARNING" THEN
        			EMWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
        			TRANSMIT
        		END IF
        	END IF
        END IF
	END IF
END IF 'confirmed address from maxis
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
	CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & original_mail_line_one)
	CALL write_variable_in_CASE_NOTE("                             " & original_mail_line_two & original_mail_city_line & ", " & original_mail_state_line & " " &   original_mail_zip_line)
 ELSE
	CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & original_resi_addr_line_one)
	CALL write_variable_in_CASE_NOTE("                               " & original_resi_addr_line_two & original_resi_addr_city & ", " & original_resi_addr_state & " " & original_resi_addr_zip)
END IF
IF homeless_addr = "Yes" Then Call write_variable_in_CASE_NOTE("* Household is homeless")
Call write_variable_in_CASE_NOTE("* Mail received reports living in: " & county_code & " County and CM 0008.06.21 - was reviewed if applicable")
IF reservation_addr = "Yes" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
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

'Checks if this is a METS case and pops up a message box with instructions if the ADDR is incorrect.
IF METS_case_number <> "" and mets_addr_correspondence = "NO" THEN MsgBox "Please update the METS ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. Change In Circumstance Process)"

IF ADDR_actions <> "no response received" THEN
	script_end_procedure_with_error_report("Success! TIKL has been set for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over he telephone or in person, the unit is not required to also report the change on a Change Report from. ")
ELSE
	script_end_procedure_with_error_report("Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter a worker comment in PEC/WCOM.")
END IF
