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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/03/2021", "Updates for HC active/pending procedure and added handling for entering PACT panel.", "MiKayla Handley")
call changelog_update("03/12/2021", "Updated handling for current address confirmation.", "MiKayla Handley")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("02/13/2020", "Updated the zip code to only allow for 5 characters.", "MiKayla Handley, Hennepin County")
call changelog_update("06/06/2019", "Initial version. Re-written per POLI/TEMP.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------

EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
call Check_for_MAXIS(false)                         'Ensuring we are not passworded out
back_to_self                                        'added to ensure we have the time to update and send the case in the background

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Running the initial dialog
BeginDialog Dialog1, 0, 0, 221, 215, "RETURNED MAIL PROCESSING"
  EditBox 60, 5, 50, 15, MAXIS_case_number
  EditBox 60, 25, 50, 15, date_received
  DropListBox 105, 65, 110, 15, "Select:"+chr(9)+"forwarding address in MN"+chr(9)+"forwarding address outside MN"+chr(9)+"no forwarding address provided"+chr(9)+"no response received", ADDR_actions
  DropListBox 100, 150, 110, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO-attempted"+chr(9)+"NO-not attempted", client_contact_dropdown 'TO DO question do we want more HERE'
  EditBox 70, 175, 145, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 115, 195, 50, 15
    CancelButton 165, 195, 50, 15
    PushButton 135, 5, 80, 15, "SNAP TE02.08.012", SNAP_POLI_TEMP_button
    PushButton 135, 25, 80, 15, "CASH TE02.08.011", CASH_POLI_TEMP_button
    PushButton 135, 45, 80, 15, "ONEsource", one_source_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 50, 10, "Date Received:"
  Text 5, 70, 85, 10, "Mail Has Been Returned:"
  GroupBox 5, 85, 210, 85, " Request for Contact"
  Text 10, 100, 140, 10, "Please make an attempt to call the client."
  Text 10, 115, 200, 10, "SNAP: shelter cost must be verified to allow shelter deduction"
  Text 10, 125, 195, 10, "CASH: address must be verified"
  Text 10, 135, 200, 15, "HC: currently all MA cases must stay open even if address is unknown"
  Text 45, 155, 50, 10, "Client Contact:"
  Text 5, 180, 60, 10, "Worker Signature:"
EndDialog

DO
    DO
    	err_msg = ""
    	DIALOG Dialog1
    	cancel_confirmation
    	IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
		IF isdate(date_received) = FALSE THEN
			err_msg = err_msg & vbnewline & "Please enter the date."
		Else
			IF Cdate(date_received) > cdate(date) = TRUE THEN err_msg = err_msg & vbnewline & "You must enter an actual date that is not in the future and is in the footer month that you are working in."
		End If
		IF ADDR_actions = "Select:" THEN err_msg = err_msg & vbCr & "Please chose an action for the returned mail."
		IF client_contact_dropdown = "Select:" THEN err_msg = err_msg & vbCr & "Please chose an action for the client contact."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
		IF ButtonPressed = CASH_POLI_TEMP_button THEN CALL view_poli_temp("02", "08", "011", "") 'TE02.08.011' RETURNED MAIL PROCESSING - CASH
		IF ButtonPressed = SNAP_POLI_TEMP_button THEN CALL view_poli_temp("02", "08", "012", "") 'TE02.08.012' RETURNED MAIL PROCESSING - SNAP
		IF ButtonPressed = one_source_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/cs/login/login.htm"
		IF ButtonPressed = one_source_button or ButtonPressed = SNAP_POLI_TEMP_button or ButtonPressed = CASH_POLI_TEMP_button THEN
			err_msg = "LOOP"
		Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		End If
	Loop until err_msg = ""
	call back_to_self ' this is for if the worker has used the POLI/TEMP navigation'
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = False					'loops until user passwords back in

'setting the footer month to make the updates in'
CALL convert_date_into_MAXIS_footer_month(date_received, MAXIS_footer_month, MAXIS_footer_year)
MAXIS_footer_month_confirmation

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, please request access to the case and run the script again.")

EMReadScreen error_message, 75, 24, 2 'reading for messages that might be missed if they are not inhibiting'
error_message = trim(error_message)

EMReadScreen worker_number, 7, 21, 21              'reading the current(primary) workers number '
IF left(worker_number, 4) <> "X127" THEN script_end_procedure("*** Out of County ***" & vbCr & worker_number & vbCr & "Please resolve for the script to continue.")

' CALL read_ADDR_panel(addr_eff_date, addr_future_date, resi_addr_line_one, resi_addr_line_two,              resi_addr_city, resi_addr_state, resi_addr_zip,                                                                                              mail_line_one, mail_line_two,                   mail_city_line, mail_state_line, mail_zip_line, living_situation, living_sit_line, homeless_line, addr_phone_1A)
Call access_ADDR_panel("READ", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_county, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to navigate'
case_invalid_error = trim(case_invalid_error)
If case_invalid_error = "ENTER A VALID COMMAND OR PF-KEY" Then case_invalid_error = ""			'this message is for if you press transmit on a single instance panel and does not indicate an error. Ignoring it
IF trim(case_invalid_error) <> "" THEN script_end_procedure("*** NOTICE!!! ***" & vbCr & case_invalid_error & vbCr & "Please resolve for the script to continue.")

CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

snap_or_cash_case = False 'we need to establish if snap or cash are active, pending, for how the case is handled'
If family_cash_case = True Then snap_or_cash_case = True
If mfip_case = True Then snap_or_cash_case = True
If dwp_case = True Then snap_or_cash_case = True
If adult_cash_case = True Then snap_or_cash_case = True
If ga_case = True Then snap_or_cash_case = True
If msa_case = True Then snap_or_cash_case = True
If grh_case = True Then snap_or_cash_case = True
If snap_case = True Then snap_or_cash_case = True

'-------------------------------------------------------------------------------------------------DIALOG
residential_address_confirmed = "YES"
If mail_street_full <> "" Then
	residential_address_confirmed = "NO"
	mailing_address_confirmed = "YES"
End If
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 566, 105, "RETURNED MAIL PROCESSING"
  DropListBox 220, 15, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", residential_address_confirmed
  DropListBox 500, 15, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", mailing_address_confirmed
  EditBox 75, 85, 205, 15, notes_on_address
  ButtonGroup ButtonPressed
    PushButton 285, 85, 50, 15, "CASE/ADHI", ADHI_button
    PushButton 355, 85, 70, 15, "HSR Manual/SNAP", HSR_manual_button
    OkButton 445, 85, 55, 15
    CancelButton 505, 85, 55, 15
  Text 10, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
  Text 15, 30, 200, 10, resi_addr_street_full
  Text 15, 40, 200, 10, resi_addr_city &  ", "  & resi_addr_state & " "   & resi_addr_zip
  Text 10, 90, 65, 10, "Notes on Address:"
  GroupBox 285, 5, 275, 75, "Mailing Address"
  Text 290, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
  Text 295, 30, 200, 10, mail_street_full
  If mail_city <> "_______________" Then Text 295, 40, 200, 10, mail_city_line & ", "  & mail_state_line &  " "  & mail_zip_line
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
			DIALOG Dialog1
			cancel_without_confirmation
			MAXIS_dialog_navigation
            IF buttonpressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/return_Mail_Processing_for_SNAP.aspx") 'HSR manual policy page
        LOOP until ButtonPressed = -1
		IF residential_address_confirmed = "Select:" THEN err_msg = err_msg & vbCr & "Please confirm if this the address that the agency attempted to deliver mail to."
		IF mail_line_one <> "" THEN
			IF mailing_address_confirmed = "Select:" THEN err_msg = err_msg & vbCr & "Please confirm if the mailing address is the address that the agency attempted to deliver mail to."
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
        BeginDialog Dialog1, 0, 0, 186, 215, "RETURNED MAIL PROCESSING NO FORWARDING ADDRESS PROVIDED"
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
        		IF verif_request_checkbox = UNCHECKED and CRF_checkbox = UNCHECKED and SVF_checkbox= UNCHECKED THEN err_msg = err_msg & vbCr & "Please select the verification requested and ensure forms are sent in ECF."
				IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
				IF mets_addr_correspondence = "YES" THEN
        			IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid METS case number."
        		END IF
        		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
        	LOOP UNTIL err_msg = ""
        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows ser     to password back into MAXIS
        LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	END IF

	return_mail_checkbox = CHECKED ' it is implied in the title of the script'
	IF ADDR_actions = "forwarding address in MN"  THEN
		new_addr_state = "MN"
		reservation_addr = "No"
		reservation_name = "N/A"
	END IF

	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	IF ADDR_actions = "forwarding address in MN" or ADDR_actions = "forwarding address outside MN" THEN
    	BeginDialog Dialog1, 0, 0, 206, 215, "RETURNED MAIL PROCESSING"
    	  CheckBox 10, 15, 75, 10, "Returned Mail/Other", return_mail_checkbox
    	  CheckBox 100, 15, 90, 10, "Shelter Form (DHS-2952)", SVF_checkbox
    	  CheckBox 10, 25, 80, 10, "Verification Request", verif_request_checkbox
    	  CheckBox 100, 25, 100, 10, "Change Report (DHS-2402)", CRF_checkbox
    	  EditBox 40, 55, 155, 15, new_addr_line_one
    	  EditBox 40, 75, 155, 15, new_addr_line_two
    	  EditBox 40, 95, 155, 15, new_addr_city
    	  EditBox 40, 115, 20, 15, new_addr_state
    	  EditBox 155, 115, 40, 15, new_addr_zip
    	  DropListBox 145, 140, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", mets_addr_correspondence
    	  EditBox 145, 155, 55, 15, mets_case_number
    	  EditBox 60, 175, 140, 15, other_notes
    	  ButtonGroup ButtonPressed
    	    OkButton 95, 195, 50, 15
    	    CancelButton 150, 195, 50, 15
    	  GroupBox 5, 5, 195, 35, "Verification Request Includes:"
    	  GroupBox 5, 40, 195, 95, "Forwarding Address:"
    	  Text 10, 60, 25, 10, "Street:"
    	  Text 10, 100, 15, 10, "City:"
    	  Text 10, 120, 20, 10, "State:"
    	  Text 120, 120, 30, 10, "Zip code:"
    	  Text 10, 145, 95, 10, "METS correspondence sent:"
    	  Text 10, 160, 70, 10, "METS case number:"
    	  Text 10, 180, 40, 10, "Other notes:"
    	EndDialog

    	DO
    		DO
    			err_msg = ""
    			DIALOG Dialog1
    			cancel_without_confirmation
				new_addr_line_one = trim(UCASE(new_addr_line_one))
				new_addr_line_two = trim(UCASE(new_addr_line_two))
				new_addr_city = trim(UCASE(new_addr_city))
				new_addr_state = trim(new_addr_state)
				new_addr_zip = trim(new_addr_zip)
    			IF new_addr_line_one = "" THEN err_msg = err_msg & vbCr & "Please complete the street address the client in now living at."
    			IF new_addr_city = "" THEN err_msg = err_msg & vbCr & "Please complete the city in which the client in now living."
    			IF new_addr_state = "" THEN err_msg = err_msg & vbCr & "Please complete the state in which the client in now living."
    			IF new_addr_zip = "" OR (new_addr_zip <> "" AND len(new_addr_zip) > 5) THEN err_msg = err_msg & vbNewLine & "Please only enter a 5 digit zip code."     'Makes sure there is a numeric zip
    			IF mets_addr_correspondence = "Select:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."
    			IF mets_addr_correspondence = "YES" THEN
    				IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
    			END IF
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    		LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to password back into MAXIS
    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		new_addr_street_full = ""

		begining_of_footer_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
		begining_of_footer_month = DateAdd("d", 0, begining_of_footer_month)
		If DateDiff("d", begining_of_footer_month, addr_eff_date) > 0 Then begining_of_footer_month = addr_eff_date
		Call access_ADDR_panel("WRITE", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, county_code, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
	END IF

    IF ADDR_actions = "no response received" THEN
        Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 206, 165, "RETURNED MAIL PROCESSING - No Response Received"
		  CheckBox 10, 15, 75, 10, "Returned Mail/Other", return_mail_checkbox
		  CheckBox 10, 25, 80, 10, "Verification Request", verif_request_checkbox
		  CheckBox 100, 15, 90, 10, "Shelter Form (DHS-2952)", SVF_checkbox
		  CheckBox 100, 25, 100, 10, "Change Report (DHS-2402)", CRF_checkbox
		  EditBox 155, 45, 45, 15, date_requested
		  DropListBox 155, 65, 45, 20, "Select:"+chr(9)+"YES"+chr(9)+"NO", ECF_reveiwed 'it they select no this is casenoted'
		  EditBox 55, 85, 145, 15, other_notes
		  ButtonGroup ButtonPressed
		    OkButton 105, 105, 45, 15
		    CancelButton 155, 105, 45, 15
		  Text 5, 50, 100, 10, "Date verification(s) requested:"
		  Text 5, 90, 45, 10, "Other Notes:"
		  Text 5, 70, 105, 10, "ECF reviewed for verifications?"
		  Text 5, 130, 195, 30, "Allow the household 10 days to respond before proceeding with a termination notice and ensure that the action is appropriate for the active programs."
		  GroupBox 5, 5, 195, 35, "Verification(s) Requested:"
		EndDialog

        DO
        	DO
        		err_msg = ""
        		DIALOG Dialog1
        		cancel_without_confirmation
        		If isdate(date_requested) = FALSE THEN  err_msg = err_msg & vbnewline & "Please enter the date verifications were requested."
				IF Cdate(date_requested) > cdate(date) = TRUE THEN  err_msg = err_msg & vbnewline & "You must enter an actual date that is not in the future."
        		IF ECF_reveiwed = "Select:" THEN  err_msg = err_msg & vbnewline & "Please review ECF to ensure the requested verifications are not on file."
	    	LOOP UNTIL err_msg = ""
        	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to      password back into MAXIS
        LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

        'per POLI/TEMP this only pertains to active cash and snap '
        IF snap_or_cash_case = TRUE THEN
        	CALL MAXIS_background_check
			Dialog1 = "" 'Running the initial dialog
			BeginDialog Dialog1, 0, 0, 351, 65, "Returned Mail Processing PACT panel entry"
			  ButtonGroup ButtonPressed
			    OkButton 240, 45, 50, 15
			    CancelButton 295, 45, 50, 15
			    PushButton 5, 45, 65, 15,  "PACT TE02.13.10", POLI_TEMP_PACT_button
			  Text 5, 5, 340, 20, "When mail to the client has been returned to the agency and the agency has sent a Request for Contact to the client with no response, for SNAP programs used code 4 will be used and code 3 for CASH."
			  Text 5, 30, 305, 15, " Allow 10 days for the client to respond to the Verification Request before terminating benefits. "
			EndDialog

			DO
        		err_msg = ""
        		DIALOG Dialog1
        		cancel_without_confirmation
				IF ButtonPressed = POLI_TEMP_PACT_button THEN CALL view_poli_temp("02", "13", "10", "") 'TE02.13.10' STAT:  PACT
				IF ButtonPressed = POLI_TEMP_PACT_button THEN
					err_msg = "LOOP"
				Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
					IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
				End If
				IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
        	LOOP UNTIL err_msg = ""
			back_to_SELF 'only need to go back if the button is pressed'
			CALL navigate_to_MAXIS_screen("STAT", "PACT") 	'Checking to see if the PACT panel is empty, if not it create a new panel'
        	EMReadScreen panel_number, 1, 02, 73
        	If panel_number = "0" then
        		EMWriteScreen "NN", 20,79 'cursor is automatically set to 06, 58'
        		TRANSMIT
        	ELSE
				PF9
			END IF
        	EMReadScreen open_cash1, 2, 6, 43 'I have to read these'
			EMReadScreen open_cash2, 2, 8, 43
			IF open_cash1 <> "  " THEN EMWriteScreen "3", 6, 58 'Enter code "3" (Refused/Failed Required Info)'
			IF open_cash2 <> "  " THEN EMWriteScreen "3", 8, 58 'Enter code "3" (Refused/Failed Required Info)'
        	If snap_case = True THEN EMWriteScreen "4", 12, 58 'Enter code "4" (Refused/Failed (FS Only))'
        	IF grh_case = True THEN EMWriteScreen "3", 10, 58 'Enter code "3" (Refused/Failed Required Info)'
        	TRANSMIT
        	EMReadScreen pop_upmsg, 7, 11, 08
        	IF pop_upmsg = "WARNING" THEN
        		EmWriteScreen "Y", 13, 64 ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
        		TRANSMIT
        	END IF

			EMReadScreen error_message,  74, 24, 02 'for script_run_lowdown'
	       	IF trim(error_message) = "" THEN
        		case_note_only = FALSE
        	else
        		maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & error_message & vbNewLine, vbYESNO + vbQuestion, "Message handling") 'this should only come up if the case is inactive'
        		IF maxis_error_check = vbNO THEN
        			case_note_only = FALSE 'this will case note only'
        		END IF
        		IF maxis_error_check= vbYES THEN
        			case_note_only = TRUE 'this will update the panels and case note'
        		END IF
        	END IF
        END IF 'if snap or cash are true'
		'we cannot close HC currently but this is the place for that handling'
	END IF 'if no response received'
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
due_date = dateadd("d", 10, date_requested)	'defaults the due date for all verifications at 10 days
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF ADDR_actions <> "no response received" THEN Call create_TIKL("Returned mail rec'd contact from the client should have occurred regarding address change. If no response-verbal or written, please take appropriate action.", 10, date, True, TIKL_note_text)
'starts a blank case note
'----------------------------------------------------------------------------------------------------CASENOTE
CALL start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("Returned mail received " & ADDR_actions)
CALL write_bullet_and_variable_in_CASE_NOTE("Received on", date_received)
CALL write_bullet_and_variable_in_CASE_NOTE("Client contacted", client_contact_dropdown)
IF mailing_address_confirmed = "YES" THEN  'Address Detail
	CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & mail_line_one)
	If mail_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                             " & mail_line_two)
	CALL write_variable_in_CASE_NOTE("                               " & mail_city_line & ", " & mail_state_line & " " &   mail_zip_line)
END IF
IF residential_address_confirmed = "YES" THEN
	CALL write_variable_in_CASE_NOTE("* Returned mail received from: " & resi_addr_line_one)
	If resi_addr_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                               " & resi_addr_line_two)
	CALL write_variable_in_CASE_NOTE("         		                      " & resi_addr_city & ", " & resi_addr_state & " " & resi_addr_zip)
	CALL write_bullet_and_variable_in_case_note("Verification(s) received", pending_verifs)
END IF
IF homeless_addr = "Yes" Then Call write_variable_in_CASE_NOTE("* Household is homeless")
IF reservation_addr = "Yes" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
Call write_bullet_and_variable_in_CASE_NOTE("Living Situation", living_situation)
Call write_bullet_and_variable_in_CASE_NOTE("Address Detail", notes_on_address)

IF ADDR_actions = "forwarding address in MN" or ADDR_actions = "forwarding address outside MN" THEN
	CALL write_variable_in_CASE_NOTE("* Forwarding address was on return mail.")
	CALL write_variable_in_CASE_NOTE("* Mailing address updated:  " & new_addr_line_one)
	If new_addr_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two)
	CALL write_variable_in_CASE_NOTE("                            " & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
	CALL write_bullet_and_variable_in_case_note("request was sent", due_date)
ELSEIF ADDR_actions = "no response received" THEN
	IF ECF_reveiwed = "YES" THEN
		CALL write_variable_in_CASE_NOTE ("* ECF reviewed for requested verifications")
		CALL write_bullet_and_variable_in_case_note("Verification(s) requested", pending_verifs)
		CALL write_bullet_and_variable_in_case_note("Verification(s) request date", date_requested)
		CALL write_bullet_and_variable_in_case_note("Due date", due_date)
	ELSE
		CALL write_variable_in_CASE_NOTE ("* ECF not reviewed for requested verifications")
	END IF
	IF snap_or_cash_case = True THEN CALL write_variable_in_CASE_NOTE ("* PACT panel entered per POLI/TEMP TE02.13.10")
	CALL write_variable_in_case_note("* Client must be provided 10 days to return requested verifications")
END IF

IF mets_addr_correspondence <> "Select:" THEN CALL write_bullet_and_variable_in_CASE_NOTE("METS correspondence sent", mets_addr_correspondence)
CALL write_bullet_and_variable_in_CASE_NOTE("METS case number", METS_case_number)
CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_run_lowdown = script_run_lowdown & vbCr & " Message: " & vbCr & error_message & vbCr & ADDR_actions & "ADDR_actions " & vbCr & "notes_on_address " & notes_on_address & vbCr & "resi address " & resi_addr_line_one & " " & resi_addr_line_two & " " & resi_addr_street_full & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip & vbCr & "resi_county " & resi_county & vbCr & "addr_verif " & addr_verif & vbCr & "homeless_addr " & homeless_addr & vbCr & "reservation_addr " & reservation_addr & vbCr & "living_situation " & living_situation & vbCr & "reservation_name " & reservation_name & vbCr & "mail address " & mail_line_one & " " & mail_line_two & " " & mail_street_full & " " & mail_city_line & " " & mail_state_line & " " & mail_zip_line & vbCr & "addr_eff_date & addr_future_date " & addr_eff_date & addr_future_date & vbCr & "phone " & phone_one & phone_two & phone_three & vbCr & "addr_email " & addr_email & vbCr & "verif_received " & verif_received & vbCr & "original_information " & original_information & vbCr & "update_attempted " & update_attempted & vbCr & "Returned Mail " & return_mail_checkbox & vbCr & "Shelter Form " & SVF_checkbox & vbCr & "Verification Request " & verif_request_checkbox & vbCr & "Change Report (DHS-2402) " & CRF_checkbox & vbCr & "new addr " & new_addr_line_one & " " & new_addr_line_two & " " & new_addr_city  & " " & new_addr_state & " " & new_addr_zip & vbCr & "county list" & county_code & vbCr & "Mets " & mets_addr_correspondence & mets_case_number & vbCr & "Other Notes " & other_notes
'Checks if this is a METS case and pops up a message box with instructions if the ADDR is incorrect.
IF METS_case_number <> "" and mets_addr_correspondence = "NO" THEN MsgBox "Please update the METS ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. Change In Circumstance Process)"

IF ADDR_actions <> "no response received" THEN
	closing_message = closing_message & vbCr & vbCr & "Success! TIKL has been set for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over the telephone or in person, the unit is not required to also report the change on a Change Report from. "  & vbCr & end_msg
ELSE
    closing_message = closing_message & vbCr & vbCr & "Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter using NOTICES SPEC/WCOM adding worker comments." & vbCr & end_msg '[this meets the requirement for HC'
END IF
Call script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step---------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/11/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/11/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/11/2022
'--All variables in dialog match mandatory fields-------------------------------06/14/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/11/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------05/11/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/14/2022
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/11/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------05/11/2022
'--PRIV Case handling reviewed -------------------------------------------------05/11/2022
'--Out-of-County handling reviewed----------------------------------------------05/11/2022	discussed with Ilse
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/11/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/11/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/14/2022
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------06/14/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/14/2022
'--comment Code-----------------------------------------------------------------06/14/2022
'--Update Changelog for release/update------------------------------------------06/14/2022
'--Remove testing message boxes-------------------------------------------------06/14/2022
'--Remove testing code/unnecessary code-----------------------------------------06/14/2022
'--Review/update SharePoint instructions----------------------------------------06/14/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/14/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/14/2022
'--Complete misc. documentation (if applicable)---------------------------------06/14/2022
'--Update project team/issue contact (if applicable)----------------------------06/14/2022
'--Other Note-------------------------------------------------------------------The reason we updated the residential instead of the mailing address is the inibiting errors that occur when updating the panel
'the last iteration this was discussed at length and the determinaition was made to leave the residential address and only update mailing although that has not been the best practive at Hennepin Co. I have sent 3 Pf11 about this clarification '
'COUNTY OF RESIDENCE CANNOT BE 89 WHEN STATE IS MN - we are not updating the residential piece'
'SNAP 2. On STAT/ADDR, enter the new address from the returned mail envelope.  Enter "OT" in the verification field. We are not updating OT as it is in the residential area of the script
