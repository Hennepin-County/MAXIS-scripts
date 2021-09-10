'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PF11 ACTIONS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
STATS_denomination = "C"       		'M is for each MEMBER
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
call changelog_update("06/04/2021", "Updated to add option for HC REVW Dates need update GitHub Issue #168.", "MiKayla Handley, Hennepin County")
call changelog_update("08/07/2020", "Updated to review CASE/NOTE for previous PF11 request.", "MiKayla Handley, Hennepin County")
call changelog_update("07/20/2019", "Per DHS Bulletin #18-69-02C, updated New Spouse Income handling and case note.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""

CALL check_for_maxis(FALSE) 'checking for passord out, brings up dialog'
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
End If
'Running the dialog for case number and client
Do
	err_msg = ""
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 201, 220, "PF11 Action"
	  EditBox 55, 5, 45, 15, MAXIS_case_number
	  DropListBox 80, 25, 115, 15, HH_Memb_DropDown, clt_to_update
	  DropListBox 80, 45, 115, 15, "Select One:"+chr(9)+"PMI Merge Request"+chr(9)+"Non-Actionable DAIL Removal"+chr(9)+"Case Note Removal Request"+chr(9)+"MFIP New Spouse Income"+chr(9)+"New Spouse Income Determination"+chr(9)+"Unable to update HC REVW dates"+chr(9)+"Other", PF11_actions
	  EditBox 80, 65, 115, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 100, 125, 45, 15
	    CancelButton 150, 125, 45, 15
	  CheckBox 15, 105, 160, 10, "Please check here if a PF11 is not required,", no_pf11_checkbox
	  Text 5, 10, 45, 10, "Case number:"
	  Text 5, 30, 70, 10, "Household member:"
	  Text 5, 50, 50, 10, "Select action:"
	  Text 5, 70, 60, 10, "Worker signature:"
	  Text 5, 85, 185, 20, "The system being down, issuance problems, or any type     of emergency should NOT be reported via a PF11."
	  Text 40, 115, 90, 10, "but a case note is needed."
	  ButtonGroup ButtonPressed
	    PushButton 110, 5, 85, 15, "HH MEMB SEARCH", search_button
	  Text 10, 160, 185, 50, "On the SELF menu and type TASK. If you have the task number enter it and it will take you directly into the PF11. If you do not have the task number or wish to look at a list of all the PF11s you have created, change the Option in TASK from T-task to a C-creator.  By placing an X next to a PF11 listed you will be able to view it."
	  GroupBox 5, 150, 190, 65, "How to check a PF11 status:"
	EndDialog


	Dialog Dialog1
	IF ButtonPressed = cancel Then StopScript
	IF ButtonPressed = search_button Then
		If MAXIS_case_number = "" Then
			MsgBox "Cannot search without a case number, please try again."
		Else
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
			err_msg = err_msg & "Start Over"
		End If
	End If
	IF MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a valid case number."
	IF PF11_actions <> "Non-Actionable DAIL Removal" THEN
	 	IF clt_to_update = "Select One:" Then err_msg = err_msg & vbNewLine & "Please select a client to update."
	END IF
	IF PF11_actions = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select the action you wish to take."
	IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "Enter your worker signature."
	IF err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

IF PF11_actions = "Non-Actionable DAIL Removal" THEN
	Call Navigate_to_MAXIS_screen ("DAIL", "DAIL")
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 316, 85, "Non-Actionable DAIL Removal"
	  EditBox 45, 5, 35, 15, message_date
	  EditBox 110, 5, 200, 15, message_to_use
	  EditBox 110, 25, 200, 15, request_reason
	  EditBox 55, 45, 255, 15, other_notes
	  EditBox 110, 65, 36, 15, worker_number
	  ButtonGroup ButtonPressed
	    OkButton 205, 65, 50, 15
	    CancelButton 260, 65, 50, 15
	  Text 5, 10, 40, 10, "DAIL Date:"
	  Text 5, 30, 105, 10, "Reason DAIL is non-actionable:"
	  Text 90, 10, 20, 10, "DAIL:"
	  Text 5, 50, 45, 10, "Other Notes:"
	  Text 5, 70, 95, 10, "Worker X127 # (last 3 digits):"
	EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "Please enter a dail date."
			IF message_to_use = "" THEN err_msg = err_msg & vbNewLine & "Please enter a dail. This can be done via copy and paste."
			IF request_reason = "" THEN err_msg = err_msg & vbNewLine & "Please enter a request reason."
			IF worker_number = len(worker_number) > 3 then err_msg = err_msg & vbNewLine & "Please enter the worker X127 number. Must be last 3 digits."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Case Note Removal Request" THEN ' this does not leave a case note'
	Call Navigate_to_MAXIS_screen ("CASE", "NOTE")
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 316, 85, "Case Note Removal Request"
      EditBox 45, 5, 35, 15, message_date
      EditBox 120, 5, 190, 15, message_to_use
      EditBox 75, 25, 235, 15, request_reason
      EditBox 50, 45, 260, 15, other_notes
      EditBox 110, 65, 40, 15, worker_number
      ButtonGroup ButtonPressed
        OkButton 205, 65, 50, 15
        CancelButton 260, 65, 50, 15
      Text 5, 10, 35, 10, "Note Date:"
      Text 5, 30, 70, 10, "Reason for removal:"
      Text 90, 10, 30, 10, "Header:"
      Text 5, 50, 45, 10, "Other Notes:"
      Text 5, 70, 95, 10, "Worker X127 # (last 3 digits):"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "Please enter the case note date."
			IF message_to_use = "" THEN err_msg = err_msg & vbNewLine & "Please enter the case note header. This can be copy and paste."
			IF request_reason = "" THEN err_msg = err_msg & vbNewLine & "Please enter a request reason."
			IF worker_number = "" or len(worker_number) > 3 then err_msg = err_msg & vbNewLine & "Please enter the worker X127 number. Must be last 3 digits."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Other" THEN
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 326, 110, "Other"
      EditBox 70, 10, 240, 15, request_reason
      EditBox 70, 30, 240, 15, other_notes
      EditBox 70, 50, 240, 15, action_taken
      EditBox 70, 70, 60, 15, worker_number
      ButtonGroup ButtonPressed
        OkButton 205, 70, 50, 15
        CancelButton 260, 70, 50, 15
      Text 5, 15, 60, 10, "Describe Problem:"
      Text 25, 35, 45, 10, "Other Notes:"
      Text 15, 55, 50, 10, "Actions Taken:"
      Text 10, 75, 55, 10, "Worker Number:"
      Text 5, 95, 230, 10, "While the dialog box is open navigate to the panel you wish to report."
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
			IF request_reason = "" THEN err_msg = err_msg & vbNewLine & "Please enter a request reason."
    		IF worker_number = "" or len(worker_number) > 3 then err_msg = err_msg & vbNewLine & "Please enter the worker X127 number. Must be last 3 digits."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

If PF11_actions = "PMI Merge Request" or PF11_actions = "MFIP New Spouse Income" or PF11_actions = "New Spouse Income Determination" THEN
	six_months_prior = DateAdd("m", -6, date) 'will set the date 6 months prior to the run date '
	'handling to prevent duplicate case notes or PF11 requests'
	Call Navigate_to_MAXIS_screen ("CASE", "NOTE")
	note_row = 5        'these always need to be reset when looking at Case note
	note_date = ""
	note_title = ""
	Do
		EMReadScreen note_date, 8, note_row, 6
		EMReadScreen note_title, 55, note_row, 25
		note_title = trim(note_title)
		IF left(note_title, 14) = "PF11 Requested" or left(note_title, 25) = "***PMI MERGE REQUESTED***" THEN
		    DO
		    	prog_confirmation = MsgBox("*** NOTICE!***" & vbNewLine & "a PF11 was already requested on: " & note_date & vbNewLine & "please do not send a duplicate request." & vbNewLine & "Select YES to continue NO to exit the script.", vbYesNo, "Possible Duplicate Request")
		    	IF prog_confirmation = vbNo THEN script_end_procedure_with_error_report("The script has ended. The request has not been acted on.")
		    	IF prog_confirmation = vbYes THEN
		    		EXIT DO
		    	END IF
		    Loop
		END IF
		IF note_date = "        " then Exit Do
		note_row = note_row + 1
		IF note_row = 19 THEN
			PF8
			note_row = 5
		END IF
		EMReadScreen next_note_date, 8, note_row, 6
		IF next_note_date = "        " then Exit Do
	Loop until datevalue(next_note_date) < six_months_prior 'looking ahead at the next case note kicking out the dates before app'

	Call Navigate_to_MAXIS_screen ("STAT", "MEMB")
	'redefine ref_numb'
	MEMB_number = left(clt_to_update, 2)	'Settin the reference number
	EMWriteScreen MEMB_number, 20, 76
	TRANSMIT
	EMReadScreen client_first_name, 12, 6, 63
	client_first_name = replace(client_first_name, "_", "")
	client_first_name = trim(client_first_name)
	EMReadScreen client_last_name, 25, 6, 30
	client_last_name = replace(client_last_name, "_", "")
	client_last_name = trim(client_last_name)
	EMReadScreen client_DOB_month, 02, 08, 42
	EMReadScreen client_DOB_date, 02, 08, 45
	EMReadScreen client_DOB_year, 04, 08, 48
END IF

If PF11_actions = "PMI Merge Request" THEN
	EMReadScreen panel_check, 4, 2, 48
	IF panel_check <> "MEMB" THEN MsgBox "*** NOTICE!***" & vbNewLine & "Case must be on STAT/MEMB to read the correct information, please re-run the script" & vbNewLine
	PF2 'going to PERS'
	EMWriteScreen client_last_name, 04, 36
	client_last_name = trim(client_last_name)
	client_last_name = replace(client_last_name, "_", "")
	'msgbox client_last_name
	EMWriteScreen client_first_name, 10, 36
	client_first_name = trim(client_first_name)
	client_first_name = replace(client_first_name, "_", "")
	'MsgBox client_first_name
	EMWriteScreen client_DOB_month, 11, 53
	EMWriteScreen client_DOB_date, 11, 56
	EMWriteScreen client_DOB_year, 11, 59
	TRANSMIT

    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 276, 135, "PMI Merge Requested"
	  EditBox 80, 5, 50, 15, PMI_number
	  EditBox 80, 25, 50, 15, PMI_number_two
	  EditBox 220, 25, 50, 15, second_case_number
	  DropListBox 80, 45, 190, 15, "Select One:"+chr(9)+"METS case opened"+chr(9)+"PMI number assigned thru SMI or PMIN"+chr(9)+"Incorrect information on case"+chr(9)+"Other", reason_request_dropdown
	  EditBox 80, 65, 190, 15, action_taken
	  EditBox 80, 85, 190, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 165, 115, 50, 15
	    CancelButton 220, 115, 50, 15
	  Text 5, 70, 50, 10, " Actions taken:"
	  Text 5, 10, 65, 10, "PMI on this case:"
	  Text 150, 5, 115, 15, "If additional PMI(s) are found add to other notes"
	  Text 5, 90, 45, 10, "Other notes:"
	  Text 5, 30, 60, 10, "Duplicate PMI(s):"
	  Text 5, 50, 65, 10, "Reason for request:"
	  Text 150, 30, 65, 10, "Other case number:"
	  Text 5, 105, 260, 10, "Review CASE/NOTE to ensure a request has not been made previously"
	EndDialog

	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			If PMI_number = "" THEN err_msg = err_msg & vbNewLine & "Please enter the PMI on this case."
			If trim(second_case_number) = "" and reason_request_dropdown <> "PMI number assigned thru SMI or PMIN" THEN err_msg = err_msg & vbNewLine & "Please enter the second case number, if none enter N/A."
			If reason_request_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the request reason."
			If reason_request_dropdown = "Other" and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please enter the request reason in other notes."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF

IF PF11_actions = "MFIP New Spouse Income" then
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 346, 140, "Designated Spouse"
  	  ButtonGroup ButtonPressed
        OkButton 235, 120, 50, 15
        CancelButton 290, 120, 50, 15
  	  Text 5, 5, 335, 20, "The Designated Spouse is the person whose income will not be counted when the policy is applied. Once determined, the Designated Spouse remains the same for the 12 consecutive months."
  	  Text 5, 30, 340, 15, "If only one newly-married member is in an existing assistance unit, the spouse joining the assistance unit will be the Designated Spouse."
  	  Text 5, 55, 340, 20, "If both newly-married members are part of the same OR different existing assistance units, they may choose, but are not required to choose, who the Designated Spouse is."
  	  Text 5, 80, 335, 35, "If the newly-married members do not choose a Designated Spouse, the eligibility worker must select the spouse with the highest combined counted income at the time. If neither spouse has income, the worker will wait to select a Designated Spouse. The worker will select the first spouse to have counted earned or unearned income during the 12 consecutive months as the Designated Spouse."
	EndDialog
	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 151, 215, "MFIP New Spouse Income"
	  EditBox 90, 5, 55, 15, new_spouse_income
	  EditBox 90, 25, 55, 15, marriage_date
	  EditBox 90, 45, 55, 15, marriage_date_verified
	  EditBox 70, 65, 75, 15, verif_used
	  EditBox 85, 85, 60, 15, total_income
	  EditBox 120, 105, 25, 15, household_size
	  EditBox 90, 125, 55, 15, worker_phone_number
	  ButtonGroup ButtonPressed
	    OkButton 50, 195, 45, 15
	    CancelButton 100, 195, 45, 15
	  Text 5, 10, 85, 10, "New Spouse Income Eff:"
	  Text 5, 30, 55, 10, "Marriage Date:"
	  Text 5, 50, 50, 10, "Date Verified:"
	  Text 5, 90, 50, 10, "Total Income:"
	  Text 5, 110, 35, 10, "HH Size:"
	  Text 5, 130, 85, 10, "Worker's Direct Contact:"
	  Text 5, 150, 140, 40, "FYI: After the Production install, MAXIS staff will contact workers to update the appropriate STAT panel(s) and ensure the case is caught by automation on an ongoing basis."
	  Text 5, 70, 65, 10, "Verification Used:"
	EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
			IF new_spouse_income = "" THEN err_msg = err_msg & vbNewLine & "Please enter the the new spouse income."
			IF marriage_date = "" THEN err_msg = err_msg & vbNewLine & "Please enter the date of marriage."
			IF marriage_date_verified = "" THEN err_msg = err_msg & vbNewLine & "Please enter the date marriage was verified."
			IF verif_used = "" THEN err_msg = err_msg & vbNewLine & "Please enter the verificaton of marriage used."
			IF total_income = "" THEN err_msg = err_msg & vbNewLine & "Please enter the total income."
			IF household_size = "" THEN err_msg = err_msg & vbNewLine & "Please enter the Household size."
			IF worker_phone_number = "" then err_msg = err_msg & vbNewLine & "Please enter the worker's phone number."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "New Spouse Income Determination" THEN
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 151, 225, "Determination New Spouse Income"
      EditBox 90, 5, 55, 15, new_spouse_income
      EditBox 90, 25, 55, 15, marriage_date
      EditBox 90, 45, 55, 15, marriage_date_verified
      EditBox 70, 65, 75, 15, verif_used
      EditBox 95, 85, 50, 15, total_income
      EditBox 120, 105, 25, 15, MEMB_number
      EditBox 120, 125, 25, 15, household_size
      EditBox 95, 145, 50, 15, income_limit
      DropListBox 95, 165, 50, 15, "Select One:"+chr(9)+"Over "+chr(9)+"Under", income_test_dropdown
      EditBox 100, 185, 45, 15, date_income_not_counted
      ButtonGroup ButtonPressed
        OkButton 50, 205, 45, 15
        CancelButton 100, 205, 45, 15
      Text 5, 10, 85, 10, "New Spouse Income Eff:"
      Text 5, 30, 55, 10, "Marriage Date:"
      Text 5, 50, 50, 10, "Date Verified:"
      Text 5, 70, 65, 10, "Verification Used:"
      Text 5, 90, 50, 10, "Total Income:"
      Text 5, 110, 110, 10, "Designated Spouse Ref Number:"
      Text 5, 130, 35, 10, "HH Size:"
      Text 5, 150, 65, 10, "275% FPG Amount: "
      Text 5, 170, 85, 10, "New Spouse Income Test:"
      Text 5, 190, 90, 10, "DS Income Not Counted Eff:"
    EndDialog
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
            IF new_spouse_income = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the the new spouse income."
            IF marriage_date = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date of marriage."
            IF marriage_date_verified = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date marriage was verified."
            IF verif_used = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the verificaton of marriage used."
            IF total_income = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the total income."
            IF MEMB_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the MEMB number."
            IF household_size = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the HH size."
            IF income_test_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please advise if the client is over or under the %275 FPG test."
            IF income_limit = "" THEN err_msg = err_msg & vbNewLine & "* Please enter if the the amount of the 275% for the HH size."
            IF date_income_not_counted = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the date the income is not counted in maxis."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Unable to update HC REVW dates" THEN
	Call Navigate_to_MAXIS_screen ("DAIL", "DAIL")
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 316, 95, "Unable to update HC REVW dates"
	  EditBox 60, 5, 35, 15, message_date
	  EditBox 185, 5, 125, 15, date_edit
	  EditBox 60, 25, 125, 15, other_notes
	  EditBox 270, 25, 40, 15, worker_number
	  ButtonGroup ButtonPressed
	    OkButton 205, 60, 50, 15
	    CancelButton 260, 60, 50, 15
	  Text 5, 10, 50, 10, "Review Date:"
	  Text 100, 10, 80, 10, "Last Review Date Edit:"
	  Text 5, 30, 45, 10, "Other Notes:"
	  Text 215, 30, 55, 10, "Worker X127 #:"
	  Text 5, 50, 185, 10, "This will send a PF11 for MNIT to update the review date"
	  Text 5, 65, 175, 10, "This information can be found directly in OneSource"
	  Text 5, 80, 235, 10, "Review SIR Email for a response when this process is complete."
	EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "Please enter a dail date."
			IF date_edit = "" THEN err_msg = err_msg & vbNewLine & "Please enter the edit you recieved when trying to change the REVW date. This can be done via copy and paste."
			IF worker_number = len(worker_number) > 3 then err_msg = err_msg & vbNewLine & "Please enter the worker X127 number. Must be last 3 digits."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions <> "New Spouse Income Determination" and no_pf11_checkbox = UNCHECKED THEN
 	PF11
	'Problem.Reporting
	EMReadScreen nav_check, 4, 1, 27
	IF nav_check = "Prob" THEN
    	IF PF11_actions = "Case Note Removal Request" or PF11_actions = "Non-Actionable DAIL Removal" or PF11_actions = "Other" or PF11_actions = "Unable to update HC REVW dates" THEN
    	    EMWriteScreen PF11_actions & " for case number: " & MAXIS_case_number, 05, 07
    	    EMWriteScreen "Date: " & message_date, 06, 07
    	    IF PF11_actions = "Case Note Removal Request" THEN EMWriteScreen "Case Note: " & message_to_use, 07, 07
    	    IF PF11_actions = "Non-Actionable DAIL Removal" THEN EMWriteScreen "DAIL: " & message_to_use, 07, 07
    		IF PF11_actions = "Non-Actionable DAIL Removal" THEN EMWriteScreen "Other error to report:", 08, 07
			IF PF11_actions = "Unable to update HC REVW dates" THEN
				EMWriteScreen "Last Review Date Edit: " & date_edit, 08, 07
				EMWriteScreen "Other notes: " & other_notes, 09, 07
			END IF
    	    IF request_reason <> "" THEN EMWriteScreen "Reason for request: " & request_reason, 09, 07
    		EMWriteScreen "Worker number: X127" & worker_number , 10, 07
    	END IF
    	IF PF11_actions = "PMI Merge Request" THEN
    		EMWriteScreen "PMI merge request for case number: " & MAXIS_case_number, 05, 07
    		EMWriteScreen "Current case PMI number: " & PMI_number, 06, 07
    		IF PMI_number_two <> "" THEN EMWriteScreen "Duplicate PMI number: " & PMI_number_two, 07, 07
    		EMWriteScreen "Reason for request: " & reason_request_dropdown, 08, 07
    		IF second_case_number <> "" THEN EMWriteScreen "Associated case number: " & second_case_number, 09, 07
    	END IF
    	If PF11_actions = "MFIP New Spouse Income" then
			EMWriteScreen PF11_actions & " for case number: " & MAXIS_case_number, 05, 07
			EMWriteScreen "Designated Spouse: " & client_first_name & " " & client_last_name, 06, 07
			EMWriteScreen "Designated Spouse Ref Number: " & MEMB_number, 07, 07
			EMWriteScreen "New Spouse Income Effective: " & new_spouse_income, 08, 07
			EMWriteScreen "Marriage Date: " & marriage_date, 09, 07
			EMWriteScreen "Marriage Date Verified: " & marriage_date_verified, 10, 07
			EMWriteScreen "Verification Used: " & verif_used, 11, 07
			EMWriteScreen "Total Income: " & total_income, 12, 07
			EMWriteScreen "HH Size: " & household_size, 13, 07
			EMWriteScreen "Worker Number: " & worker_number, 14, 07
    	END IF
    	TRANSMIT
    	EMReadScreen task_number, 7, 3, 27
    	TRANSMIT
    	PF3 ''-self'
    	PF3 '- MEMB'
    ELSE
		script_end_procedure_with_error_report("Could not reach PF11." & PF11_actions & " has not been sent.")
	END IF
END IF

reminder_date = dateadd("d", 5, date)
Call change_date_to_soonest_working_day(reminder_date, "BACK")

IF PF11_actions = "PMI Merge Request" or PF11_actions = "Non-Actionable DAIL Removal" or PF11_actions = "MFIP New Spouse Income" or PF11_actions = "New Spouse Income Determination" or PF11_actions = "Other" or PF11_actions = "Unable to update HC REVW dates" THEN
	CALL start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	IF PF11_actions = "PMI Merge Request"  THEN CALL write_variable_in_case_note("PF11 Requested " & PF11_actions & " for M" & MEMB_number)
	IF PF11_actions = "Non-Actionable DAIL Removal" or PF11_actions = "Other" or PF11_actions = "Unable to update HC REVW dates"  THEN CALL write_variable_in_case_note("PF11 Requested " & PF11_actions)
	IF PF11_actions = "Non-Actionable DAIL Removal" or PF11_actions = "Other" or PF11_actions = "PMI Merge Request" or PF11_actions = "Unable to update HC REVW dates"  THEN
	    CALL write_bullet_and_variable_in_CASE_NOTE("Reason for request", reason_request_dropdown)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Reason for request", request_reason)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Task number", task_number)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Requested for", client_info)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Message", message_to_use)
		CALL write_bullet_and_variable_in_CASE_NOTE("Review Date", message_date)
		CALL write_bullet_and_variable_in_CASE_NOTE("Last Review Date Edit", date_edit)
		IF PF11_actions = "Unable to update HC REVW dates"  THEN
			CALL write_variable_in_CASE_NOTE("Sent a PF11 for MNIT to update the review date")
			CALL write_variable_in_CASE_NOTE("This information can be found directly in OneSource")
			CALL write_variable_in_CASE_NOTE("SIR Email will be reviewed for a response when this process is complete.")
		END IF
	    If PMI_number <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("PMI number", PMI_number)
	    If PMI_number_two <> "" then Call write_bullet_and_variable_in_CASE_NOTE("Duplicate PMI number", PMI_number_two)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Associated case number", second_case_number)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	    call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
	    CALL write_variable_in_CASE_NOTE ("---")
	    CALL write_variable_in_CASE_NOTE(worker_signature)
	    PF3
	END IF
	IF  PF11_actions = "MFIP New Spouse Income" and no_pf11_checkbox = UNCHECKED THEN
		CALL write_variable_in_case_note(PF11_actions & " for M" & MEMB_number)
		CALL write_bullet_and_variable_in_CASE_NOTE("PF11 Sent - Task number", task_number)
		CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", message_date)
	    CALL write_bullet_and_variable_in_CASE_NOTE("New Spouse Income Effective", new_spouse_income)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", marriage_date)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Date Marriage Was Verified", marriage_date_verified)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Used", verif_used)
	    CALL write_bullet_and_variable_in_CASE_NOTE("Total Income", total_income)
	    CALL write_bullet_and_variable_in_CASE_NOTE("HH Size", household_size)
	    CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG Amount", income_limit)
	    CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG New Spouse Income Test", income_test_dropdown)
	    CALL write_bullet_and_variable_in_CASE_NOTE("DS income not counted eff", date_income_not_counted)
		call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
		CALL write_variable_in_CASE_NOTE ("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
		PF3
	END IF

	IF  PF11_actions = "MFIP New Spouse Income" and no_pf11_checkbox = CHECKED THEN
		CALL write_variable_in_case_note(PF11_actions & " for M" & MEMB_number)
		CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", message_date)
		CALL write_bullet_and_variable_in_CASE_NOTE("New Spouse Income Effective", new_spouse_income)
		CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", marriage_date)
		CALL write_bullet_and_variable_in_CASE_NOTE("Date Marriage Was Verified", marriage_date_verified)
		CALL write_bullet_and_variable_in_CASE_NOTE("Verification Used", verif_used)
		CALL write_bullet_and_variable_in_CASE_NOTE("Total Income", total_income)
		CALL write_bullet_and_variable_in_CASE_NOTE("HH Size", household_size)
		CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG Amount", income_limit)
		CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG New Spouse Income Test", income_test_dropdown)
		CALL write_bullet_and_variable_in_CASE_NOTE("DS income not counted eff", date_income_not_counted)
		call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
		CALL write_variable_in_CASE_NOTE ("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
		PF3
	END IF
    'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
    Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "PF11 check: " & PF11_actions & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
END IF

IF PF11_actions = "New Spouse Income Determination" THEN
	CALL start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	CALL write_variable_in_case_note(PF11_actions & " for M" & MEMB_number) 	''---Determination MFIP New Spouse Income for M19---
	CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", marriage_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("New Spouse Income Effective", new_spouse_income)
    CALL write_bullet_and_variable_in_CASE_NOTE("Marriage Date", marriage_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Date Marriage Was Verified", marriage_date_verified)
    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Used", verif_used)
    CALL write_bullet_and_variable_in_CASE_NOTE("Total Income", total_income)
    CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG HH Size", household_size)
    CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG Amount", income_limit)
    CALL write_bullet_and_variable_in_CASE_NOTE("275% FPG New Spouse Income Test", income_test_dropdown)
    CALL write_bullet_and_variable_in_CASE_NOTE("DS income not counted eff", date_income_not_counted)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3
END IF

script_end_procedure_with_error_report("It worked! " & PF11_actions & " has been sent and/or case noted.")
