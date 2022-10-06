''GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-MATCH CLEARED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine & "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine & "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
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

script_run_lowdown = ""

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/06/2022", "Update to remove hard coded DEU signature all DEU scripts.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/28/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/08/2020", "Updated BUG when clearing match BO-Other added back IULB notes per DEU request.", "MiKayla Handley, Hennepin County") '#922
CALL changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley, Hennepin County") '#493
CALL changelog_update("08/24/2021", "Remove mandatory handling from other notes variable.", "MiKayla Handley, Hennepin County") '#571 '
CALL changelog_update("06/09/2021", "Handling for script end procedure.", "MiKayla Handley, Hennepin County") '#373 '
CALL changelog_update("01/11/2021", "Updated BNDX handling to ensure header of case note is written correctly.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/20/2020", "Removed custom functions from script file. Functions have all been incorporated into the project's Function Library.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/17/2020", "The field for 'OTHER NOTES' is now required when completing the information to clear the match. ##~## ##~##We are aware that this will not always be required in MAXIS and will be adding additional functionality for scenario and match specific requirements of this field, but in order to provide you with a working script right now this field must be mandatory each time.##~## ##~##Thank you for your patience as we provide updates to this script.##~##", "Casey Love, Hennepin County")
CALL changelog_update("09/08/2020", "Updated BUG when clearing match BO-Other worker must indicate other notes for comments on IULA.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/22/2020", "Combined OP script with match cleared, added HH member dialog. Created a new drop down for claim referral tracking.", "MiKayla Handley, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
CALL changelog_update("07/17/2019", "Updated script to no longer run off DAIL, it will ask for a case number to ensure all the matches pull correctly.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/14/2019", "Updated dialog and case note to reflect BE-Child requirements.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/23/2018", "Updated case note to reflect standard dialog and case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/26/2018", "Merged the claim referral tracking back into the script.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/16/2018", "Corrected case note for pulling IEVS period.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match when the date is over 45 days.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match BE-OP entered.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/13/2017", "Updated correct handling for BEER matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/08/2017", "Now includes handling for sending the difference notice and clearing the WAGE match including NC codes.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/27/2017", "Added BP - Wrong Person", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/22/2017", "Updated Non-coop option to the cleared match.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/21/2017", "Updated to clear match, and added handling for sending the difference notice.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/14/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
CALL check_for_maxis(FALSE) 'checking for password out, brings up dialog'
CALL MAXIS_case_number_finder(MAXIS_case_number)

If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
End If
'Running the initial dialog to confirm what type match is being cleared and the specifics about the case
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
DO
	DO
	   err_msg = ""
       Dialog1 = ""
       BeginDialog Dialog1, 0, 0, 201, 85, "Match Cleared"
         EditBox 55, 5, 45, 15, MAXIS_case_number
         DropListBox 80, 25, 115, 15, HH_Memb_DropDown, clt_to_update
         EditBox 80, 45, 115, 15, worker_signature
         ButtonGroup ButtonPressed
           PushButton 110, 5, 85, 15, "HH MEMB SEARCH", search_button
           OkButton 90, 65, 50, 15
           CancelButton 145, 65, 50, 15
         Text 5, 30, 70, 10, "Household member:"
         Text 5, 50, 60, 10, "Worker signature:"
         Text 5, 10, 45, 10, "Case number:"
       EndDialog

	    Dialog Dialog1
	    cancel_confirmation
	    Call validate_MAXIS_case_number(err_msg, "*")
	    IF ButtonPressed = search_button Then 'this will check for if the worker is on the DAIL and the script cant find a case number'
	    	IF MAXIS_case_number = "" Then
	    		MsgBox "Cannot search without a case number, please try again."
	    	Else
	    		HH_Memb_DropDown = ""
	    		Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
	    		err_msg = err_msg & "Start Over"
	    	End If
	    End If
		IF clt_to_update = "Select One:" Then err_msg = err_msg & vbNewLine & "Please select a client to update."
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""						'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

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
EMReadscreen client_mid_initial, 1, 6, 79
EMReadScreen client_DOB, 10, 8, 42
EMReadscreen client_SSN, 11, 7, 42
client_SSN = replace(client_SSN, " ", "")

'navigating to INFC
CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(client_SSN, 3, 63)
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
EMReadScreen panel_check, 4, 2, 52
IF panel_check <> "IEVP" THEN script_end_procedure_with_error_report("***NOTICE***" & vbNewLine & "Case must be on INFC/IEVP to read the correct information. If the social security number is not found the match must be completed manually. The only way to find the wage match is go to REPT/IEVC. The issue might be that the client has a duplicate PMI number. Review for a PF11 to be submitted.")

'------------------------------------------------------------------selecting the correct wage match
Row = 7
DO
	EMReadScreen IEVS_period, 11, row, 47
	EMReadScreen number_IEVS_type, 3, row, 41
	IF trim(IEVS_period) = "" THEN script_end_procedure_with_error_report("A match for the selected period could not be found. The script will now end.")
	BeginDialog Dialog1, 0, 0, 171, 95, "CASE NUMBER: "  & MAXIS_case_number
  	 Text 5, 10, 100, 10, "Navigate to the correct match:"
  	 Text 5, 25, 150, 10, "Match Type: " & number_IEVS_type
  	 Text 5, 40, 150, 10, "Match Period: "  & IEVS_period
  	 ButtonGroup ButtonPressed
     PushButton 5, 60, 50, 15, "Confirm Match", match_confimation
     PushButton 60, 60, 50, 15, "Next Match", next_match
     PushButton 115, 60, 50, 15, "Next Page", next_page
    CancelButton 60, 80, 50, 15
	EndDialog
	DO
	    DO
	       	err_msg = ""
	       	Dialog Dialog1
			cancel_confirmation
			IF ButtonPressed = next_match THEN
				row = row + 1
				IF row = 17 THEN
					PF8
					row = 7
					EMReadScreen IEVS_period, 11, row, 47
				END IF
			END IF
			IF ButtonPressed = next_page THEN
				PF8
				row = 7
				EMReadScreen IEVS_period, 11, row, 47
			END IF
			IF ButtonPressed = match_confimation THEN EXIT DO
	        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	       LOOP UNTIL err_msg = ""
		CALL check_for_password_without_transmit(are_we_passworded_out)
	LOOP UNTIL are_we_passworded_out = false
LOOP UNTIL ButtonPressed = match_confimation

'---------------------------------------------------------------------Reading potential errors for out-of-county cases
CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" THEN
	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
ELSE
    EMReadScreen number_IEVS_type, 3, 7, 12 'read the match type'
    IF number_IEVS_type = "A30" THEN match_type = "BNDX"
    IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
    IF number_IEVS_type = "A70" THEN match_type = "BEER"
    IF number_IEVS_type = "A80" THEN match_type = "UNVI"
    IF number_IEVS_type = "A60" THEN match_type = "UBEN"
    IF number_IEVS_type = "A50" THEN match_type = "WAGE"
		IF number_IEVS_type = "A51" THEN match_type = "WAGE"
	IEVS_year = ""
	IF match_type = "WAGE" THEN
		EMReadScreen select_quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
	ELSEIF match_type = "UBEN" THEN
		EMReadScreen IEVS_month, 2, 5, 68
		EMReadScreen IEVS_year, 4, 8, 71
	ELSEIF match_type = "BEER" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
	ELSEIF match_type = "UNVI" THEN
		EMReadScreen IEVS_year, 4, 8, 15
		select_quarter = "YEAR"
	END IF
END IF

'--------------------------------------------------------------------Client name
EMReadScreen panel_name, 4, 02, 52
IF panel_name <> "IULA" THEN script_end_procedure_with_error_report("Script did not find IULA.")
EMReadScreen client_name, 35, 5, 24
client_name = trim(client_name)                         'trimming the client name
IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
	length = len(client_name)                           'establishing the length of the variable
	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
ELSEIF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
ELSE                                'In cases where the last name takes up the entire space, THEN the client name becomes the last name
	first_name = ""
	last_name = client_name
END IF
first_name = trim(first_name)
IF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
END IF

'----------------------------------------------------------------------------------------------------ACTIVE PROGRAMS
EMReadScreen Active_Programs, 13, 6, 68
Active_Programs = trim(Active_Programs)
programs = ""
IF instr(Active_Programs, "D") THEN programs = programs & "DWP, "
IF instr(Active_Programs, "F") THEN programs = programs & "Food Support, "
IF instr(Active_Programs, "H") THEN programs = programs & "Health Care, "
IF instr(Active_Programs, "M") THEN programs = programs & "Medical Assistance, "
IF instr(Active_Programs, "S") THEN programs = programs & "MFIP, "
'trims excess spaces of programs
programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
'----------------------------------------------------------------------------------------------------Employer info & difference notice info
IF match_type = "UBEN" THEN income_source = "Unemployment"
IF match_type = "UNVI" THEN income_source = "NON-WAGE"
IF match_type = "WAGE" THEN
    EMReadScreen income_source, 50, 8, 37 'was 37' should be to the right of employer and the left of amount
    income_source = trim(income_source)
    length = len(income_source)		'establishing the length of the variable
    'should be to the right of employer and the left of amount '
    IF instr(income_source, " AMOUNT: $") THEN
	    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
	    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
        position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
        income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	END IF
END IF
IF match_type = "BEER" THEN
    EMReadScreen income_source, 50, 8, 28 'was 37' should be to the right of employer and the left of amount
	income_source = trim(income_source)
	length = len(income_source)		'establishing the length of the variable
	'should be to the right of employer and the left of amount '
    IF instr(income_source, " AMOUNT: $") THEN
	    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
	    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
        position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
        income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	END IF
END IF

'----------------------------------------------------------------------------------------------------notice sent
EMReadScreen notice_sent, 1, 14, 37
EMReadScreen sent_date, 8, 14, 68
sent_date = trim(sent_date)
IF sent_date = "" THEN sent_date = "N/A"
IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")
EMReadScreen clear_code, 2, 12, 58
'----------------------------------------------------------------Defaulting checkboxes to being checked (per DEU instruction)
IF notice_sent = "N" THEN
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 271, 185, "DIFFERENCE NOTICE NOT SENT FOR: " & MAXIS_case_number
	  DropListBox 85, 90, 70, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", difference_notice_action_dropdown
	  CheckBox 175, 15, 70, 10, "Difference Notice", diff_notice_checkbox
	  CheckBox 175, 25, 90, 10, "Authorization to Release", ATR_verf_checkbox
	  CheckBox 175, 35, 90, 10, "Employment Verification", EVF_checkbox
	  CheckBox 175, 45, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
	  CheckBox 175, 55, 80, 10, "Rental Income Form", rental_checkbox
	  CheckBox 175, 65, 80, 10, "Other (please specify)", other_checkbox
	  CheckBox 10, 170, 115, 10, "Set a TIKL due to 10 day cutoff", tenday_checkbox
	  DropListBox 145, 120, 115, 15, "Select One:"+chr(9)+"Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
	  EditBox 50, 145, 215, 15, other_notes
	  Text 5, 10, 165, 10, "Client name: "   & client_name
	  Text 5, 55, 160, 10, "Active Programs: "  & programs
	  Text 5, 70, 165, 15, "Income source:   " & income_source
	  ButtonGroup ButtonPressed
	    OkButton 180, 165, 40, 15
	    CancelButton 225, 165, 40, 15
	  Text 5, 25, 150, 10, "Match Type: " & match_type
	  Text 5, 40, 150, 10, "Match Period: " & IEVS_period
	  GroupBox 170, 5, 95, 75, "Verification(s) Requested: "
	  GroupBox 5, 110, 260, 30, "SNAP or MFIP Federal Food only"
	  Text 10, 125, 130, 10, "Claim Referral Tracking on STAT/MISC:"
	  Text 5, 95, 80, 10, "Send Difference Notice: "
	  Text 5, 150, 40, 10, "Other notes:"
	EndDialog

	DO
    	err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	IF difference_notice_action_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select an answer to continue."
		IF claim_referral_tracking_dropdown =  "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select if the claim referral tracking needs to be updated."
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please ensure you are completing other notes"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out) 'this cannot have a trasnmit due to navigation in IUL screens'
END IF

IF difference_notice_action_dropdown =  "YES" THEN '--------------------------------------------------------------------sending the notice in IULA
    EMwritescreen "005", 12, 46 'writing the resolve time to read for later
    EMwritescreen "Y", 14, 37 'send Notice
	TRANSMIT 'goes into IULA
	'removed the IULB information '
	TRANSMIT'exiting IULA, helps prevent errors when going to the case note
    '-----------------------------------------------------------------------------------Claim Referral Tracking
    action_date = date & ""
ELSEIF notice_sent = "Y" or difference_notice_action_dropdown =  "NO" THEN 'or clear_code <> "__" '
	'-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 170, "MATCH CLEARED - CASE NUMBER: "  & MAXIS_case_number
      EditBox 175, 5, 15, 15, resolve_time
      DropListBox 75, 35, 115, 15, "Select One:"+chr(9)+"CB-Ovrpmt And Future Save"+chr(9)+"CC-Overpayment Only"+chr(9)+"CF-Future Save"+chr(9)+"CA-Excess Assets"+chr(9)+"CI-Benefit Increase"+chr(9)+"CP-Applicant Only Savings"+chr(9)+"BC-Case Closed"+chr(9)+"BE-Child"+chr(9)+"BE-No Change"+chr(9)+"BE-NC-Non-collectible"+chr(9)+"BE-Overpayment Entered"+chr(9)+"BN-Already Known-No Savings"+chr(9)+"BI-Interface Prob"+chr(9)+"BO-Other"+chr(9)+"BP-Wrong Person"+chr(9)+"BU-Unable To Verify"+chr(9)+"NC-Non Cooperation", resolution_status
      DropListBox 120, 50, 70, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", change_response
      DropListBox 120, 65, 70, 15, "Select One:"+chr(9)+"DISQ Added"+chr(9)+"DISQ Deleted"+chr(9)+"Pending Verif"+chr(9)+"No"+chr(9)+"N/A", DISQ_action
      EditBox 275, 15, 40, 15, date_received
      CheckBox 200, 30, 70, 10, "Difference Notice", diff_notice_checkbox
      CheckBox 200, 40, 90, 10, "Authorization to Release", ATR_verf_checkBox
      CheckBox 200, 50, 90, 10, "Employment verification", EVF_checkbox
      CheckBox 200, 60, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
      CheckBox 200, 70, 80, 10, "Rental Income Form", rental_checkbox
      CheckBox 200, 80, 80, 10, "Other (please specify)", other_checkbox
      EditBox 275, 95, 40, 15, exp_grad_date
      CheckBox 5, 85, 115, 10, "Set a TIKL due to 10 day cutoff", tenday_checkbox
      CheckBox 5, 100, 130, 10, "Overpayment (other programs)", HC_OP_checkbox
      DropListBox 140, 125, 175, 15, "Select One:"+chr(9)+"Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
      EditBox 50, 150, 180, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 235, 150, 40, 15
        CancelButton 280, 150, 40, 15
      Text 5, 10, 100, 10, "Match Type: " & match_type
      Text 5, 25, 185, 10, "Match Period: " & IEVS_period
      Text 110, 10, 65, 10, "Resolve time (min): "
      GroupBox 195, 5, 125, 110, "Verification Used to Clear: "
      GroupBox 5, 115, 315, 30, "SNAP or MFIP Federal Food only"
      Text 10, 130, 130, 10, "Claim Referral Tracking on STAT/MISC:"
      Text 5, 40, 60, 10, "Resolution Status: "
      Text 5, 55, 110, 10, "Responded to Difference Notice: "
      Text 5, 70, 75, 10, "DISQ panel addressed:"
      Text 5, 155, 40, 10, "Other notes: "
      Text 200, 20, 75, 10, "Date verif rcvd/on file:"
      Text 200, 100, 65, 10, "Expected grad date:"
    EndDialog

	DO
		err_msg = ""
		DIALOG Dialog1
		cancel_without_confirmation
		other_notes = trim(other_notes)
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "Please enter a valid numeric resolved time, ie 005."
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise what other verification was used to clear the match."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF resolution_status = "BE-No Change" AND other_notes = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		IF resolution_status = "BE-Child" AND exp_grad_date = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE - Child graduation date and date rcvd must be completed."
		If resolution_status = "CC-Overpayment Only" AND programs = "Health Care" or programs = "Medical Assistance" THEN err_msg = err_msg & vbNewLine & "System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		If resolution_status = "BO-Other" AND other_notes = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BO-Other other notes must be completed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)

	IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
	    discovery_date = date
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
	    BeginDialog Dialog1, 0, 0, 361, 260, "MATCH CLEARED - CASE NUMBER: "  & MAXIS_case_number
		  Text 5, 5, 245, 15, "Income source: " & income_source
		  DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	      EditBox 65, 25, 40, 15, discovery_date
	      DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
	      EditBox 130, 65, 30, 15, OP_from
	      EditBox 180, 65, 30, 15, OP_to
	      EditBox 245, 65, 35, 15, Claim_number
	      EditBox 305, 65, 45, 15, Claim_amount
	      DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
	      EditBox 130, 85, 30, 15, OP_from_II
	      EditBox 180, 85, 30, 15, OP_to_II
	      EditBox 245, 85, 35, 15, Claim_number_II
	      EditBox 305, 85, 45, 15, Claim_amount_II
	      DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
	      EditBox 130, 105, 30, 15, OP_from_III
	      EditBox 180, 105, 30, 15, OP_to_III
	      EditBox 245, 105, 35, 15, claim_number_III
	      EditBox 305, 105, 45, 15, Claim_amount_III
	      DropListBox 50, 125, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_IV
	      EditBox 130, 125, 30, 15, OP_from_IV
	      EditBox 180, 125, 30, 15, OP_to_IV
	      EditBox 245, 125, 35, 15, claim_number_IV
	      EditBox 305, 125, 45, 15, Claim_amount_IV
		  EditBox 130, 155, 30, 15, HC_from
		  EditBox 180, 155, 30, 15, HC_to
		  EditBox 245, 155, 35, 15, HC_claim_number
		  EditBox 305, 155, 45, 15, HC_claim_amount
		  EditBox 80, 155, 20, 15, HC_resp_memb
		  EditBox 305, 175, 45, 15, Fed_HC_AMT
	      CheckBox 235, 205, 120, 10, "Earned income disregard allowed", EI_checkbox
	      EditBox 70, 200, 160, 15, EVF_used
	      EditBox 200, 25, 45, 15, income_rcvd_date
	      EditBox 70, 220, 285, 15, Reason_OP
		  EditBox 330, 25, 20, 15, OT_resp_memb
		  CheckBox 70, 240, 105, 10, "EVF/ATR is still needed", ATR_needed_checkbox
		  ButtonGroup ButtonPressed
		    OkButton 260, 240, 45, 15
		    CancelButton 310, 240, 45, 15
		  Text 265, 30, 60, 10, "OT resp. Memb #:"
		  Text 260, 10, 50, 10, "Fraud referral:"
		  Text 5, 30, 55, 10, "Discovery date: "
		  Text 5, 205, 65, 10, "Income verif used:"
		  Text 10, 160, 70, 10, "OT resp. Memb(s) #:"
		  Text 230, 180, 75, 10, "Total federal HC AMT:"
		  Text 5, 225, 60, 10, "Reason for Claim:"
		  Text 140, 30, 60, 10, "Date income rcvd: "
		  Text 285, 160, 20, 10, "AMT:"
		  Text 105, 160, 20, 10, "From:"
		  Text 215, 160, 25, 10, "Claim #"
		  Text 165, 160, 10, 10, "To:"
		  GroupBox 5, 145, 350, 50, "HC Programs Only"
		  Text 15, 70, 30, 10, "Program:"
		  Text 165, 70, 10, 10, "To:"
		  GroupBox 5, 45, 350, 100, "Overpayment Information"
		  Text 130, 55, 30, 10, "(MM/YY)"
	      Text 180, 55, 30, 10, "(MM/YY)"
		  Text 15, 70, 30, 10, "Program:"
	      Text 15, 110, 30, 10, "Program:"
	      Text 15, 90, 30, 10, "Program:"
		  Text 15, 130, 30, 10, "Program:"
		  Text 105, 70, 20, 10, "From:"
		  Text 105, 90, 20, 10, "From:"
		  Text 105, 110, 20, 10, "From:"
	      Text 105, 130, 20, 10, "From:"
		  Text 165, 70, 10, 10, "To:"
		  Text 165, 90, 10, 10, "To:"
		  Text 165, 110, 10, 10, "To:"
	      Text 165, 130, 10, 10, "To:"
		  Text 215, 70, 25, 10, "Claim #"
		  Text 215, 90, 25, 10, "Claim #"
		  Text 215, 110, 25, 10, "Claim #"
	      Text 215, 130, 25, 10, "Claim #"
		  Text 285, 70, 20, 10, "AMT:"
		  Text 285, 90, 20, 10, "AMT:"
		  Text 285, 110, 20, 10, "AMT:"
	      Text 285, 130, 20, 10, "AMT:"
		EndDialog
	    Do
	        Do
	        	err_msg = ""
	        	DIALOG Dialog1
	        	cancel_confirmation
	        	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	        	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
				If OP_program = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the overpayment program for the first claim."
				IF OP_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
				IF Claim_number = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
				IF Claim_amount = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	           	IF OP_program_II <> "Select:" THEN
	    			IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred II."
	        		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	        		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	        	END IF
	    		IF OP_program_III <> "Select:" THEN
	    			IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred III."
	    			IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	    			IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	    		END IF
	    		IF OP_program_IV <> "Select:" THEN
	    			IF OP_from_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred IV."
	    			IF Claim_number_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	    			IF Claim_amount_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	    		END IF
	        	IF HC_claim_number <> "" THEN
	        		IF HC_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment started."
	        		IF HC_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment ended."
	        		IF HC_claim_amount = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	        	END IF
	        	IF EVF_used = "" THEN err_msg = err_msg & vbNewLine & "* Please enter verification used for the income received. If no verification was received enter N/A."
	        	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	        LOOP UNTIL err_msg = ""
	        CALL check_for_password_without_transmit(are_we_passworded_out)
	    Loop until are_we_passworded_out = false
	END IF

	IF resolution_status = "CF-Future Save" THEN
	    Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 161, 120, "Cleared CF Future Savings"
  		DropListBox 65, 5, 90, 15, "Select One:"+chr(9)+"Case Became Ineligible"+chr(9)+"Person Removed"+chr(9)+"Benefit Increased"+chr(9)+"Benefit Decreased", IULB_result_dropdown
  		DropListBox 65, 25, 90, 15, "Select One:"+chr(9)+"One Time Only"+chr(9)+"Per Month For Nbr of Months", IULB_method_dropdown
    	EditBox 115, 40, 40, 15, IULB_savings_amount
    	EditBox 125, 60, 15, 15, IULB_start_month
    	EditBox 140, 60, 15, 15, IULB_start_year
    	EditBox 140, 80, 15, 15, IULB_months
    	ButtonGroup ButtonPressed
    	OkButton 60, 100, 45, 15
    	CancelButton 110, 100, 45, 15
    	Text 5, 10, 60, 10, "Results for IULB:"
    	Text 5, 30, 55, 10, "Method for IULB:"
    	Text 55, 45, 55, 10, "Savings Amount:"
    	Text 95, 65, 25, 10, "MM/YY"
    	Text 55, 65, 35, 10, "Start Date:"
    	Text 55, 85, 70, 10, "Months for Method R:"
		EndDialog

	    DO
	    	err_msg = ""
	    	DIALOG Dialog1
	    	cancel_confirmation
	    	IF IsNumeric(IULB_savings_amount) = false THEN err_msg = err_msg & vbNewLine & "Please enter a valid numeric amount no decimal."
	    	IF IULB_result_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB result."
	    	IF IULB_method_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB method."
			IF IULB_result_dropdown <> "Person Removed" and IULB_months <> "" THEN err_msg = err_msg & vbNewLine & "SAVINGS MONTHS NOT ALLOWED WITH MONTH CODE O"
	    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	    LOOP UNTIL err_msg = ""
	    CALL check_for_password_without_transmit(are_we_passworded_out)
 	    IF IULB_result_dropdown = "Case Became Ineligible" THEN IULB_result = "I"
	    IF IULB_result_dropdown = "Person Removed" THEN IULB_result = "R"
	    IF IULB_result_dropdown = "Benefit Increased" THEN IULB_result = "P"
	    IF IULB_result_dropdown = "Benefit Decreased" THEN IULB_result = "N"
		IF IULB_method_dropdown = "One Time Only" THEN IULB_method = "O"
		IF IULB_method_dropdown = "Per Month For Nbr of Months" THEN IULB_method = "O"
	END IF

	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMReadScreen panel_name, 4, 02, 52
	IF panel_name <> "IULA" THEN
		EMReadScreen back_panel_name, 4, 2, 52
		If back_panel_name <> "IEVP" Then
			CALL back_to_SELF
			CALL navigate_to_MAXIS_screen("INFC" , "____")
			CALL write_value_and_transmit("IEVP", 20, 71)
			CALL write_value_and_transmit(client_SSN, 3, 63) 'do not need to do the non-disclosure here '
		End If
		CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
	End If

	EMWriteScreen resolve_time, 12, 46	    'resolved notes depending on the resolution_status
    IF resolution_status = "CB-Ovrpmt And Future Save" THEN IULA_res_status = "CB"
    IF resolution_status = "CC-Overpayment Only" THEN IULA_res_status = "CC" 'Claim Entered" CC cannot be used - ACTION ODE FOR ACTH OR ACTM IS INVALID
    IF resolution_status = "CF-Future Save" THEN IULA_res_status = "CF"
    IF resolution_status = "CA-Excess Assets" THEN IULA_res_status = "CA"
    IF resolution_status = "CI-Benefit Increase" THEN IULA_res_status = "CI"
    IF resolution_status = "CP-Applicant Only Savings" THEN IULA_res_status = "CP"
    IF resolution_status = "BC-Case Closed" THEN IULA_res_status = "BC"
    IF resolution_status = "BE-Child" THEN IULA_res_status = "BE"
    IF resolution_status = "BE-No Change" THEN IULA_res_status = "BE"
    IF resolution_status = "BE-Overpayment Entered" THEN IULA_res_status = "BE"
    IF resolution_status = "BE-NC-Non-collectible" THEN IULA_res_status = "BE"
    IF resolution_status = "BI-Interface Prob" THEN IULA_res_status = "BI"
    IF resolution_status = "BN-Already Known-No Savings" THEN IULA_res_status = "BN"
    IF resolution_status = "BP-Wrong Person" THEN IULA_res_status = "BP"
    IF resolution_status = "BU-Unable To Verify" THEN IULA_res_status = "BU"
    IF resolution_status = "BO-Other" THEN IULA_res_status = "BO"
    IF resolution_status = "NC-Non Cooperation" THEN IULA_res_status = "NC"
    'checked these all to programS'

    EMwritescreen IULA_res_status, 12, 58
    IF IULA_res_status = "CC" THEN
        col = 57
        Do
        	EMReadscreen action_header, 4, 11, col
        	If action_header <> "    " Then
        		If action_header = "ACTH" Then
        			EMWriteScreen "BE", 12, col+1
				Else
        			EMWriteScreen "CC", 12, col+1
				End If
        	End If
        	col = col + 6
        Loop until action_header = "    "
    END IF

    IF change_response = "YES" THEN
    	EMwritescreen "Y", 15, 37
    ELSE
    	EMwritescreen "N", 15, 37
    END IF

    TRANSMIT 'Going to IULB
    '----------------------------------------------------------------------------------------writing the note on IULB
	other_notes = trim(other_notes)
	EMReadScreen panel_name, 4, 02, 52
    IF panel_name = "IULB" and (difference_notice_action_dropdown = "NO" OR notice_sent = "Y") THEN
    	TRANSMIT
    	EMReadScreen MISC_error_check,  74, 24, 02
    	EMReadScreen IULB_enter_msg, 5, 24, 02
    	IF IULB_enter_msg = "ENTER" OR IULB_enter_msg = "ACTIO" THEN 'check if we need to input other notes
			CALL clear_line_of_text(8, 6)
			CALL clear_line_of_text(9, 6)
			IF resolution_status = "CB-Ovrpmt And Future Save" THEN other_notes = "OP Claim entered and future savings. " & other_notes
			IF resolution_status = "CC-Overpayment Only" Or HC_OP_checkbox = CHECKED THEN
				other_notes = "Claim entered. See case note. " & other_notes
				CALL clear_line_of_text(17, 9)
			END IF
			EMWriteScreen Claim_number, 17, 9
			EMWriteScreen Claim_number_II, 18, 9
			EMWriteScreen claim_number_III, 19, 9

			IF resolution_status = "CF-Future Save" THEN
				other_notes = "Future Savings. " & other_notes
				EMwritescreen active_programs, 12, 37
				EMwritescreen IULB_results, 12, 42
				EMwritescreen IULB_method, 12, 49
				EMwritescreen IULB_savings_amount, 12, 54
				EMwritescreen IULB_start_month, 12, 65
				EMwritescreen IULB_start_year, 12, 68
				EMwritescreen IULB_months, 12, 74
				TRANSMIT
			END IF
			IF resolution_status = "CA-Excess Assets" THEN IULB_notes = "Excess Assets. " & other_notes
			IF resolution_status = "CB-Ovrpmt And Future Save" THEN IULB_notes = "OP Claim entered and future savings."
			IF resolution_status = "CC-Overpayment Only" THEN IULB_notes = "Claim entered. See case note. "
			IF resolution_status = "CI-Benefit Increase" THEN IULB_notes = "Benefit Increase. " & other_notes
			IF resolution_status = "CP-Applicant Only Savings" THEN IULB_notes = "Applicant Only Savings. " & other_notes
			IF resolution_status = "BC-Case Closed" THEN IULB_notes = "Case closed. " & other_notes
			IF resolution_status = "BE-Child" THEN IULB_notes = "No change, minor child income excluded. " & other_notes
			IF resolution_status = "BE-No Change" THEN IULB_notes = "No change. " & IULB_notes
			IF resolution_status = "BE-Overpayment Entered" THEN IULB_notes = "OP entered other programs. " & other_notes
			IF resolution_status = "BE-NC-Non-collectible" THEN IULB_notes = "Non-Coop remains, but claim is non-collectible. " & other_notes
			IF resolution_status = "BI-Interface Prob" THEN IULB_notes = "Interface Problem. " & other_notes
			IF resolution_status = "BN-Already Known-No Savings" THEN IULB_notes = "Already known - No savings. " & other_notes
			IF resolution_status = "BP-Wrong Person" THEN IULB_notes = "Client name and wage earner name are different. " & other_notes
			IF resolution_status = "BU-Unable To Verify" THEN IULB_notes = "Unable To Verify. " & other_notes
			IF resolution_status = "BO-Other" THEN IULB_notes = "No review due during the match period. " & other_notes
			IF resolution_status = "NC-Non Cooperation" THEN IULB_notes = "Non-coop, requested verf not in ECF, " & other_notes

			iulb_row = 8
			iulb_col = 6
			notes_array = split(IULB_notes, " ") 'this is where we write to IULB'
			For each word in notes_array
				EMWriteScreen word & " ", iulb_row, iulb_col
				If iulb_col + len(word) > 77 Then
					iulb_col = 6
					iulb_row = iulb_row + 1
					If iulb_row = 10 Then Exit For
				End If
				iulb_col = iulb_col + len(word) + 1
			Next
    	    TRANSMIT
    		EMReadScreen MISC_error_check,  74, 24, 02
    		IF trim(MISC_error_check) <> "" THEN
    			next_steps_message_box = MsgBox("***WARNING MESSAGE***" & vbNewLine & "Do you want to transmit?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion,     "Message handling")
    			IF next_steps_message_box = vbYes THEN
    				TRANSMIT
    				EMReadScreen panel_name, 4, 02, 52
    			END IF
    			IF next_steps_message_box= vbNo THEN
    				PF3
    				EMReadScreen panel_name, 4, 02, 52
					script_run_lowdown = script_run_lowdown & vbCr & vbCR & "DEU Error Type: " & MISC_error_check & panel_name
    			END IF
    		END IF
    	ELSE
    		CALL back_to_SELF
    	END IF
    ELSE
    	script_run_lowdown = script_run_lowdown & vbCr & vbCR & "DEU Error Type: " & MISC_error_check & panel_name
    END IF
END IF 'end of match when difference_notice_action_dropdown =  "YES" '

script_run_lowdown = script_run_lowdown & vbCr & vbCr & "Notice Sent: " & notice_sent
script_run_lowdown = script_run_lowdown & vbCr & "Sent Date: " & sent_date
script_run_lowdown = script_run_lowdown & vbCr & "DIFF NOTC ACTION: " & difference_notice_action_dropdown
script_run_lowdown = script_run_lowdown & vbCr & "Claim referral tracking: " & claim_referral_tracking_dropdown
script_run_lowdown = script_run_lowdown & vbCr & "Client Name: " & client_name
script_run_lowdown = script_run_lowdown & vbCr & "The Programs: " & programs
script_run_lowdown = script_run_lowdown & vbCr & "Income Source: " & income_source
script_run_lowdown = script_run_lowdown & vbCr & "Match Type: " & match_type
script_run_lowdown = script_run_lowdown & vbCr & "IEVS Period: " & IEVS_period
script_run_lowdown = script_run_lowdown & vbCr & "Resolve Time: " & resolve_time
script_run_lowdown = script_run_lowdown & vbCr & "Resolution Status: " & resolution_status
script_run_lowdown = script_run_lowdown & vbCr & "Change Response: " & change_response
script_run_lowdown = script_run_lowdown & vbCr & "DISQ Action: " & DISQ_action & vbCR
script_run_lowdown = script_run_lowdown & vbCr & "Active Programs Codes: " & Active_Programs
script_run_lowdown = script_run_lowdown & vbCr & "IULA Resolution Status: " & IULA_res_status
script_run_lowdown = script_run_lowdown & vbCr & "IULB Enter Msg: " & IULB_enter_msg
script_run_lowdown = script_run_lowdown & vbCr & "Other Notes: " & other_notes & vbCr
script_run_lowdown = script_run_lowdown & vbCr & "Fraud referral: " & fraud_referral
script_run_lowdown = script_run_lowdown & vbCr & "Discovery Date: " & discovery_date & vbCR
script_run_lowdown = script_run_lowdown & vbCr & "CLAIM I" & vbCR & "OP Program: " & OP_program
script_run_lowdown = script_run_lowdown & vbCr & "OP from: " & OP_from
script_run_lowdown = script_run_lowdown & vbCr & "OP to: " & OP_to
script_run_lowdown = script_run_lowdown & vbCr & "Claim Number: " & Claim_number
script_run_lowdown = script_run_lowdown & vbCr & "Claim Amount: " & Claim_amount& vbCR
script_run_lowdown = script_run_lowdown & vbCr & "CLAIM II" & vbCR & "OP Program: " & OP_program_II
script_run_lowdown = script_run_lowdown & vbCr & "OP from: " & OP_from_II
script_run_lowdown = script_run_lowdown & vbCr & "OP to: " & OP_to_II
script_run_lowdown = script_run_lowdown & vbCr & "Claim Number: " & Claim_number_II
script_run_lowdown = script_run_lowdown & vbCr & "Claim Amount: " & Claim_amount_II& vbCR
script_run_lowdown = script_run_lowdown & vbCr & "CLAIM III" & vbCR & "OP Program: " & OP_program_III
script_run_lowdown = script_run_lowdown & vbCr & "OP from: " & OP_from_III
script_run_lowdown = script_run_lowdown & vbCr & "OP to: " & OP_to_III
script_run_lowdown = script_run_lowdown & vbCr & "Claim Number: " & claim_number_III
script_run_lowdown = script_run_lowdown & vbCr & "Claim Amount: " & Claim_amount_III& vbCR
script_run_lowdown = script_run_lowdown & vbCr & "CLAIM IV" & vbCR & "OP Program: " & OP_program_IV
script_run_lowdown = script_run_lowdown & vbCr & "OP from: " & OP_from_IV
script_run_lowdown = script_run_lowdown & vbCr & "OP to: " & OP_to_IV
script_run_lowdown = script_run_lowdown & vbCr & "Claim Number: " & claim_number_IV
script_run_lowdown = script_run_lowdown & vbCr & "Claim Amount: " & Claim_amount_IV& vbCR
script_run_lowdown = script_run_lowdown & vbCr & "HC CLAIM" & vbCR & "OP from: " & HC_from
script_run_lowdown = script_run_lowdown & vbCr & "OP to: " & HC_to
script_run_lowdown = script_run_lowdown & vbCr & "Claim Number: " & HC_claim_number
script_run_lowdown = script_run_lowdown & vbCr & "Claim Amount: " & HC_claim_amount
script_run_lowdown = script_run_lowdown & vbCr & "HC Resp member: " & HC_resp_memb
script_run_lowdown = script_run_lowdown & vbCr & "FED HC Amount: " & Fed_HC_AMT
script_run_lowdown = script_run_lowdown & vbCr & "" & EI_checkbox
If EI_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Earned income disregard allowed"
script_run_lowdown = script_run_lowdown & vbCr & "EVF Used: " & EVF_used
script_run_lowdown = script_run_lowdown & vbCr & "Income Received Date: " & income_rcvd_date
script_run_lowdown = script_run_lowdown & vbCr & "OP Reason: " & Reason_OP
script_run_lowdown = script_run_lowdown & vbCr & "Other resp members: " & OT_resp_memb
If ATR_needed_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "EVF/ATR is still needed"

'-------------------------------------------------------------------The case note & case note related code
verification_needed = ""
IF Diff_Notice_Checkbox = CHECKED THEN verification_needed = verification_needed & "Difference Notice, "
IF EVF_checkbox = CHECKED THEN verification_needed = verification_needed & "EVF, "
IF ATR_Verf_CheckBox = CHECKED THEN verification_needed = verification_needed & "ATR, "
IF lottery_verf_checkbox = CHECKED THEN verification_needed = verification_needed & "Lottery/Gaming Form, "
IF rental_checkbox =  CHECKED THEN verification_needed = verification_needed & "Rental Income Form, "
IF other_checkbox = CHECKED THEN verification_needed = verification_needed & "Other, "

verification_needed = trim(verification_needed) 	'takes the last comma off of verification_needed when autofilled into dialog if more more than one app date is found and additional app is selected
IF right(verification_needed, 1) = "," THEN verification_needed = left(verification_needed, len(verification_needed) - 1)
'------------------------------------------------------------------STAT/MISC for claim referral tracking
IF claim_referral_tracking_dropdown <> "Not Needed" THEN
    'Going to the MISC panel to add claim referral tracking information
	CALL navigate_to_MAXIS_screen ("STAT", "MISC")
	Row = 6
	EMReadScreen panel_number, 1, 02, 73
	If panel_number = "0" THEN
		EMWriteScreen "NN", 20,79
		TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
		EMReadScreen MISC_error_check,  74, 24, 02
		IF trim(MISC_error_check) = "" THEN
			case_note_only = FALSE
		else
			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion, "Message handling")
			IF maxis_error_check = vbYes THEN case_note_only = TRUE 'this will case note only'
		    IF maxis_error_check= vbNo THEN case_note_only = FALSE 'this will update the panels and case note'
        End if
	END IF
END IF

Do
	'Checking to see if the MISC panel is empty, if not it will find a new line'
	EMReadScreen MISC_description, 25, row, 30
	MISC_description = replace(MISC_description, "_", "")
	If trim(MISC_description) = "" THEN
		'PF9
		EXIT DO
	Else
		row = row + 1
	End if
Loop Until row = 17
If row = 17 THEN MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")

'writing in the action taken and date to the MISC panel
PF9
'_________________________ 25 characters to write on MISC
IF claim_referral_tracking_dropdown =  "Initial" THEN MISC_action_taken = "Claim Referral Initial"
IF claim_referral_tracking_dropdown =  "OP Non-Collectible (please specify)" THEN MISC_action_taken = "Determination-Non-Collect"
IF claim_referral_tracking_dropdown =  "No Savings/Overpayment" THEN MISC_action_taken = "Determination-No Savings"
IF claim_referral_tracking_dropdown =  "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
EMWriteScreen MISC_action_taken, Row, 30
EMWriteScreen date, Row, 66
TRANSMIT
'------------------------------------------setting up case note header'
IF ATR_needed_checkbox = CHECKED THEN
	header_note = "ATR/EVF STILL REQUIRED"
ELSEIF difference_notice_action_dropdown = "YES" THEN
	cleared_header = "DIFF NOTICE SENT"
	sent_date = date
ELSEIF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
	cleared_header = "CLEARED CLAIM ENTERED "
ELSEIF resolution_status = "NC-Non Cooperation" THEN
		cleared_header = "NON-COOPERATION "
ELSEIF resolution_status <> "CC-Overpayment Only" OR resolution_status <> "NC-Non Cooperation" THEN
	cleared_header = "CLEARED " & IULA_res_status
ELSEIF resolution_status = "BE-NC-Non-collectible" THEN
	cleared_header = "CLEARED " & IULA_res_status & "Non-Collectible"
END IF

IF match_type = "BEER" THEN match_type_letter = "B"
IF match_type = "UBEN" THEN match_type_letter = "U"
IF match_type = "UNVI" THEN match_type_letter = "U"

IF match_type = "WAGE" THEN
	IF select_quarter = 1 THEN IEVS_quarter = "1ST"
	IF select_quarter = 2 THEN IEVS_quarter = "2ND"
	IF select_quarter = 3 THEN IEVS_quarter = "3RD"
	IF select_quarter = 4 THEN IEVS_quarter = "4TH"
END IF

IEVS_period = trim(IEVS_period)
IF match_type <> "UBEN" THEN IEVS_period = replace(IEVS_period, "/", " to ")
IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")
Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

'-------------------------------------------------------------------------------------------------The case note
IF claim_referral_tracking_dropdown <> "Not Needed" THEN
    start_a_blank_case_note
    IF claim_referral_tracking_dropdown =  "Initial" THEN
		CALL write_variable_in_case_note("Claim Referral Tracking - Initial")
	ELSE
		CALL write_variable_in_case_note("Claim Referral Tracking - " & MISC_action_taken)
	END IF
    CALL write_bullet_and_variable_in_case_note("Action Date", action_date)
    CALL write_bullet_and_variable_in_case_note("Active Program(s)", programs)
    CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
    CALL write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    IF case_note_only = TRUE THEN CALL write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
    CALL write_variable_in_case_note("-----")
    CALL write_variable_in_case_note(worker_signature)
END IF
start_a_blank_case_note
IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") " & cleared_header & header_note & "-----")
IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
IF match_type = "BNDX" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
CALL write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
IF  difference_notice_action_dropdown = "YES" THEN
	CALL write_bullet_and_variable_in_case_note("Verifications Requested", verification_needed)
	CALL write_variable_in_case_note("* Client must be provided 10 days to return requested verifications")
ELSE
	CALL write_bullet_and_variable_in_case_note("Verifications Received", verification_needed)
END IF
IF change_response <> "N/A" THEN CALL write_bullet_and_variable_in_case_note("Responded to Difference Notice", change_response)
IF DISQ_action <> "Select One:" THEN CALL write_bullet_and_variable_in_case_note("STAT/DISQ addressed for each program", DISQ_action)
CALL write_bullet_and_variable_in_case_note("Date verification received in ECF", date_received)
IF resolution_status = "CB-Ovrpmt And Future Save" THEN CALL write_variable_in_case_note("* OP Claim entered and future savings.")
IF resolution_status = "CF-Future Save" THEN CALL write_variable_in_case_note("* Future Savings.")
IF resolution_status = "CA-Excess Assets" THEN CALL write_variable_in_case_note("* Excess Assets.")
IF resolution_status = "CI-Benefit Increase" THEN CALL write_variable_in_case_note("* Benefit Increase.")
IF resolution_status = "CP-Applicant Only Savings" THEN CALL write_variable_in_case_note("* Applicant Only Savings.")
IF resolution_status = "BC-Case Closed" THEN CALL write_variable_in_case_note("* Case closed.")
IF resolution_status = "BE-Child" THEN
	CALL write_variable_in_case_note("* Income is excluded for minor child in school.")
	CALL write_bullet_and_variable_in_case_note("Expected graduation date", exp_grad_date)
END IF
IF resolution_status = "BE-No Change" THEN CALL write_variable_in_case_note("* No Overpayments or savings were found related to this match.")
IF resolution_status = "BE-Overpayment Entered" THEN CALL write_variable_in_case_note("* Overpayments or savings were found related to this match.")
IF resolution_status = "BE-NC-Non-collectible" THEN CALL write_variable_in_case_note("* No collectible overpayments or savings were found related to this match. Client is still non-coop.")
IF resolution_status = "BI-Interface Prob" THEN CALL write_variable_in_case_note("* Interface Problem.")
IF resolution_status = "BN-Already Known-No Savings" THEN CALL write_variable_in_case_note("* Client reported income. Correct income is in JOBS/BUSI and budgeted.")
IF resolution_status = "BP-Wrong Person" THEN CALL write_variable_in_case_note("* Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
IF resolution_status = "BU-Unable To Verify" THEN CALL write_variable_in_case_note("* Unable to verify, due to:")
IF resolution_status = "BO-Other" THEN CALL write_variable_in_case_note("* No review due during the match period.  Per DHS, reporting requirements are waived during pandemic.")
IF resolution_status = "NC-Non Cooperation" THEN
	CALL write_variable_in_case_note("* Client failed to cooperate with wage match.")
	CALL write_variable_in_case_note("* Case approved to close.")
	CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice.")
END IF
IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
    CALL write_variable_in_case_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
    IF OP_program_II <> "Select:" THEN CALL write_variable_in_case_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
    IF OP_program_III <> "Select:" THEN CALL write_variable_in_case_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" THEN CALL write_variable_in_case_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
    IF HC_claim_number <> "" THEN
    	CALL write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
    	CALL write_bullet_and_variable_in_case_note("Health Care responsible members", HC_resp_memb)
    	CALL write_bullet_and_variable_in_case_note("Total Federal Health Care amount", Fed_HC_AMT)
    	CALL write_variable_in_case_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
    IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
    CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
    CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
    CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
    CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
END IF
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_case_note("----- ----- ----- ----- -----")
CALL write_variable_in_case_note(worker_signature)

IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN '-----------------------------------------------------------------------------------------OP CASENOTE
    IF HC_claim_number <> "" THEN
    	EMWriteScreen "x", 5, 3
    	TRANSMIT
    	note_row = 4			'Beginning of the case notes
    	Do 						'Read each line
    		EMReadScreen note_line, 76, note_row, 3
    		note_line = trim(note_line)
    		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
    		message_array = message_array & note_line & vbcr		'putting the lines together
    		note_row = note_row + 1
    		If note_row = 18 THEN 									'End of a single page of the case note
    			EMReadScreen next_page, 7, note_row, 3
    			If next_page = "More: +" Then 						'This indicates there is another page of the case note
    				PF8												'goes to the next line and resets the row to read'\
    				note_row = 4
    			End If
    		End If
    	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
    	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
    	CALL create_outlook_email("HSPH.FAA.Unit.AR.Spaulding@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & discovery_date & "HC Claim # " & HC_claim_number, "CASE NOTE" & vbcr & message_array,"", False)
    END IF
	'-----------------------------------------------------------------writing the CCOL case note'
    CALL navigate_to_MAXIS_screen("CCOL", "CLSM")
    EMWriteScreen Claim_number, 4, 9
    TRANSMIT
    PF4
    EMReadScreen existing_case_note, 1, 5, 6
    IF existing_case_note = "" THEN
    	PF4
    ELSE
    	PF9
    END IF

    IF match_type = "WAGE" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
    IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
    IF match_type = "UBEN" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
    CALL write_bullet_and_variable_in_CCOL_NOTE("Discovery date", discovery_date)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Period", IEVS_period)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Active Programs", programs)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Source of income", income_source)
    CALL write_variable_in_CCOL_note("----- ----- ----- ----- ----- ----- -----")
    CALL write_variable_in_CCOL_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
    IF OP_program_II <> "Select:" THEN CALL write_variable_in_CCOL_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
    IF OP_program_III <> "Select:" THEN CALL write_variable_in_CCOL_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" THEN CALL write_variable_in_CCOL_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
    IF HC_claim_number <> "" THEN
    	CALL write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
    	CALL write_bullet_and_variable_in_CCOL_NOTE("Health Care responsible members", HC_resp_memb)
    	CALL write_bullet_and_variable_in_CCOL_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
    	CALL write_variable_in_CCOL_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Allowed")
    IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Not Allowed")
    CALL write_bullet_and_variable_in_CCOL_NOTE("Fraud referral made", fraud_referral)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Income verification received", EVF_used)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Date verification received", income_rcvd_date)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Reason for overpayment", Reason_OP)
    CALL write_bullet_and_variable_in_CCOL_NOTE("Other responsible member(s)", OT_resp_memb)
    CALL write_variable_in_CCOL_note("----- ----- ----- ----- ----- ----- -----")
    CALL write_variable_in_CCOL_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
    PF3 'to save CCOL casenote'

	'-------------------------------The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	IF tenday_checkbox = CHECKED THEN CALL create_TIKL("Unable to close due to 10 day cutoff. Verification of match should have returned by now. If not received and processed, take appropriate action.", 0, date, True, TIKL_note_text)
	script_end_procedure_with_error_report("Match has been acted on. Please take any additional action needed for your case.")
END IF

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/24/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/24/2022
'--All variables in dialog match mandatory fields-------------------------------06/24/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------06/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/24/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------06/24/2022
'--PRIV Case handling reviewed -------------------------------------------------06/24/2022
'--Out-of-County handling reviewed----------------------------------------------06/24/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/24/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/24/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------06/24/2022
'--Script name reviewed---------------------------------------------------------06/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------------------------N/A
'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/24/2022
'--comment Code-----------------------------------------------------------------06/24/2022
'--Update Changelog for release/update------------------------------------------06/24/2022
'--Remove testing message boxes-------------------------------------------------06/24/2022
'--Remove testing code/unnecessary code-----------------------------------------06/24/2022
'--Review/update SharePoint instructions----------------------------------------06/24/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/24/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/24/2022
'--Complete misc. documentation (if applicable)---------------------------------06/24/2022
'--Update project team/issue contact (if applicable)----------------------------06/24/2022
'TODO I need error proofing in multiple places on this script in and out of IULA and IULB ensuring the case and on CCOL'
'need to check about adding for multiple claims'
