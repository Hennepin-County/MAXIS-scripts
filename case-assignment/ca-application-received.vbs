'Required for statistical purposes==========================================================================================
name_of_script = "CA-APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 145                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
                                            vbOKonly + vbNewLineitical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("02/21/2018", "Added on demand waiver handling.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/16/2018", "Added case transfer confirmation coding.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/03/2017", "Email functionality - only expedited emails will be sent to Triagers.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/25/2017", "Email functionality - will create email, and send for all CASH and FS applications.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/12/2017", "Email functionality will create email, but not send it. Staff will need to send email after reviewing email.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog initial_dialog, 0, 0, 186, 110, "Application Received"
  EditBox 60, 5, 45, 15, MAXIS_case_number
  CheckBox 15, 40, 70, 10, "Not Active (APPL)", Not_Active_checkbox
  CheckBox 15, 55, 85, 10, "Active (add a program)", Active_checkbox
  CheckBox 15, 75, 125, 10, "Check if client is applying for SNAP", SNAP_checkbox
  ButtonGroup ButtonPressed
    OkButton 75, 90, 50, 15
    CancelButton 130, 90, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
  GroupBox 5, 25, 175, 45, "Was client active in MAXIS at time of application?"
EndDialog

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog initial_dialog
		IF buttonpressed = 0 THEN stopscript
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF Active_checkbox = UNCHECKED and Not_Active_checkbox = UNCHECKED then err_msg = err_msg & vbNewLine & "* Please select if the client is active or not."
		IF Active_checkbox = CHECKED and Not_Active_checkbox = CHECKED then err_msg = err_msg & vbNewLine & "* Please select only box if the client is active or not."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'-------------------------------------------------------------------DIALOG NO snap?'
'Gathers Date of application and creates MAXIS friendly dates to be sure to navigate to the correct time frame
'This only functions if case is in PND2 status
CALL navigate_to_MAXIS_screen("REPT","PND2")
dateofapp_row = 1
dateofapp_col = 1
EMSearch MAXIS_case_number, dateofapp_row, dateofapp_col
EMReadScreen MAXIS_case_name,  20, dateofapp_row, 16
EMReadScreen MAXIS_footer_month, 2, dateofapp_row, 38
EMReadScreen app_day, 2, dateofapp_row, 41
EMReadScreen MAXIS_footer_year, 2, dateofapp_row, 44
application_date = MAXIS_footer_month & "/" & app_day & "/" & MAXIS_footer_year

'If case is not in PND2 status this defaults the date information to current date to allow correct navigation
IF application_date = "  /  /  " THEN
	application_date = date
	CALL convert_date_into_MAXIS_footer_month (date, MAXIS_footer_month, MAXIS_footer_year)
END IF

'Determines which programs are currently pending in the month of application
CALL navigate_to_MAXIS_screen("STAT","PROG")

EMReadScreen err_msg, 7, 24, 02
IF err_msg = "BENEFIT" THEN	script_end_procedure ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen emer_pend, 4, 8, 74
EMReadScreen grh_pend, 4, 9, 74
EMReadScreen fs_pend, 4, 10, 74
EMReadScreen ive_pend, 4, 11, 74
EMReadScreen hc_pend, 4, 12, 74
EMReadScreen cca_pend, 4, 14, 74

'Assigns a value so the programs pending will show up in check boxes
IF cash1_pend = "PEND" THEN
	cash1_pend = CHECKED
ELSE
	cash1_pend = UNCHECKED
END IF

IF cash2_pend = "PEND" THEN
	cash2_pend = CHECKED
ELSE
	cash2_pend = UNCHECKED
END If

IF cash1_pend = CHECKED OR cash2_pend = CHECKED THEN cash_pend = CHECKED

IF emer_pend = "PEND" THEN
	emer_pend = CHECKED
ELSE
	emer_pend = UNCHECKED
END IF

IF grh_pend = "PEND" THEN
	grh_pend = CHECKED
ELSE
	grh_pend = UNCHECKED
END IF

IF fs_pend = "PEND" THEN
	fs_pend = CHECKED
ELSE
	fs_pend = UNCHECKED
END IF

IF ive_pend = "PEND" THEN
	ive_pend = CHECKED
ELSE
	ive_pend = UNCHECKED
END IF

IF hc_pend = "PEND" THEN
	hc_pend = CHECKED
ELSE
	hc_pend = UNCHECKED
END IF

IF cca_pend = "PEND" THEN
	cca_pend = CHECKED
ELSE
	cca_pend = UNCHECKED
END IF

'Defaults the date pended to today
pended_date = date & ""

IF fs_pend = CHECKED OR cash_pend = CHECKED OR grh_pend = CHECKED THEN send_appt_ltr = TRUE
'----------------------------------------------------------------------------------------------------dialogs

BeginDialog appl_detail_dialog, 0, 0, 296, 145, "APPLICATION RECEIVED"
  DropListBox 80, 5, 65, 15, "Select One:"+chr(9)+"Online"+chr(9)+"Mail"+chr(9)+"Fax", how_app_rcvd
  EditBox 230, 5, 60, 15, application_date
  DropListBox 80, 25, 65, 15, "Select One:"+chr(9)+"Addendum"+chr(9)+"ApplyMN"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Pop"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer", app_type
  EditBox 230, 25, 60, 15, confirmation_number
  CheckBox 80, 50, 30, 10, "Cash", cash_pend
  CheckBox 110, 50, 25, 10, "CCA", cca_pend
  CheckBox 140, 50, 50, 10, "Emergency", emer_pend
  CheckBox 195, 50, 30, 10, "GRH", grh_pend
  CheckBox 230, 50, 20, 10, "HC", hc_pend
  If SNAP_checkbox = CHECKED THEN CheckBox 260, 50, 30, 10, "SNAP", fs_pend 'need this to equal checked'
  EditBox 110, 65, 25, 15, worker_number
  EditBox 110, 85, 25, 15, team_number
  EditBox 50, 105, 240, 15, entered_notes
  ButtonGroup ButtonPressed
    OkButton 185, 125, 50, 15
    CancelButton 240, 125, 50, 15
  Text 5, 30, 65, 10, "Type of Application:"
  Text 160, 30, 50, 10, "Confirmation #"
  Text 160, 10, 65, 10, "Date of Application:"
  Text 5, 50, 70, 10, "Programs Applied for:"
  Text 5, 70, 100, 10, "Transfer to (last 3 digit of X#):"
  Text 5, 90, 90, 10, "Assigned to (3 digit team #):"
  Text 5, 110, 45, 10, "Other Notes:"
  Text 5, 10, 70, 10, "Application Received:"
  Text 145, 80, 145, 10, "* Script will transfer case to assigned worker"
EndDialog

BeginDialog add_detail_dialog, 0, 0, 281, 110, "ADD A PROGRAM"
  DropListBox 80, 5, 65, 15, "Select One:"+chr(9)+"Online"+chr(9)+"Mail"+chr(9)+"Fax", how_app_rcvd
  EditBox 215, 5, 60, 15, application_date
  DropListBox 80, 25, 65, 15, "Select One:"+chr(9)+"Addendum"+chr(9)+"ApplyMN"+chr(9)+"CAF"+chr(9)+"6696"+chr(9)+"HCAPP"+chr(9)+"HC-Certain Pop"+chr(9)+"LTC"+chr(9)+"MHCP B/C Cancer", app_type
  EditBox 215, 25, 60, 15, confirmation_number
  CheckBox 50, 50, 30, 10, "Cash", cash_pend
  CheckBox 85, 50, 25, 10, "CCA", cca_pend
  CheckBox 120, 50, 50, 10, "Emergency", emer_pend
  CheckBox 175, 50, 30, 10, "GRH", grh_pend
  CheckBox 210, 50, 20, 10, "HC", hc_pend
  CheckBox 245, 50, 30, 10, "SNAP", fs_pend
  EditBox 50, 70, 225, 15, entered_notes
  ButtonGroup ButtonPressed
    OkButton 170, 90, 50, 15
    CancelButton 225, 90, 50, 15
  Text 5, 30, 65, 10, "Type of Application:"
  Text 165, 30, 50, 10, "Confirmation #:"
  Text 150, 10, 65, 10, "Date of Application:"
  Text 5, 50, 40, 10, "Applied for:"
  Text 5, 75, 45, 10, "Other Notes:"
  Text 5, 10, 70, 10, "Application Received:"
EndDialog


'possible to navigate to the geocoder here...need handling for priv cases xfer must be the last step'
'------------------------------------------------------------------------------------DIALOG APPL
IF Not_Active_checkbox = CHECKED THEN
'Runs the second dialog - which gathers information about the application
    Do
    	Do
			err_msg = ""
    		Dialog appl_detail_dialog
    		cancel_confirmation
    		IF how_app_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
			IF app_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
    		IF isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid application date."
			IF worker_number = "" OR len(worker_number) <> 3 then err_msg = err_msg & vbNewLine & "* You must enter the worker number of the worker if you would like the case to be transfered by the script."
    		IF team_number = "" OR len(team_number) <> 3 then err_msg = err_msg & vbNewLine & "* You must enter the team number of the worker if you would like the case to be transfered by the script."
    		IF app_type = "ApplyMN" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "If an ApplyMN was received, you must enter the confirmation number and time received"
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in
End if
'-------------------------------------------------------------------DIALOG Addendum'
IF Active_checkbox = CHECKED THEN
	Do
    	Do
			err_msg = ""
    		Dialog add_detail_dialog
    		cancel_confirmation
			IF how_app_rcvd = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter how the application was received to the agency."
			IF app_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please enter the type of application received."
    		IF app_type = "ApplyMN" AND isnumeric(confirmation_number) = FALSE THEN err_msg = err_msg & vbNewLine & "If an ApplyMN was received, you must enter the confirmation number and time received"
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'Creates a variable that lists all the programs pending.
programs_applied_for = ""
IF cash_pend = CHECKED THEN programs_applied_for = programs_applied_for & "Cash, "
IF emer_pend = CHECKED THEN programs_applied_for = programs_applied_for & "Emergency, "
IF grh_pend = CHECKED THEN programs_applied_for = programs_applied_for & "GRH, "
IF fs_pend = CHECKED THEN programs_applied_for = programs_applied_for & "SNAP, "
IF ive_pend = CHECKED THEN programs_applied_for = programs_applied_for & "IV-E, "
IF hc_pend = CHECKED THEN programs_applied_for = programs_applied_for & "HC, "
IF cca_pend = CHECKED THEN programs_applied_for = programs_applied_for & "CCA"


'trims excess spaces of programs_applied_for
programs_applied_for = trim(programs_applied_for)
'takes the last comma off of programs_applied_for when autofilled into dialog if more more than one app date is found and additional app is selected
If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

'--------------------------------------------------------------------------------initial case note
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE ("~ Application Received (" & app_type & ") via " & how_app_rcvd & " on " & application_date & " ~")
IF Active_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE ("Client is currently active in MAXIS")
IF Not_Active_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE ("Intake")
IF isnumeric(confirmation_number) = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE ("Confirmation # ", confirmation_number)
IF app_type = "6696" THEN write_variable_in_CASE_NOTE ("Form Rcvd: MNsure Application for Health Coverage and Help Paying Costs (DHS-6696) ")
IF app_type = "HCAPP" THEN write_variable_in_CASE_NOTE ("Form Rcvd: Health Care Application (HCAPP) (DHS-3417) ")
IF app_type = "HC-Certain Pop" THEN write_variable_in_CASE_NOTE ("Form Rcvd: MHC Programs Application for Certain Populations (DHS-3876) ")
IF app_type = "LTC" THEN write_variable_in_CASE_NOTE ("Form Rcvd: Application for Medical Assistance for Long Term Care Services (DHS-3531) ")
IF app_type = "MHCP B/C Cancer" THEN write_variable_in_CASE_NOTE ("Form Rcvd: Minnesota Health Care Programs Application and Renewal Form Medical Assistance for Women with Breast or Cervical Cancer (DHS-3525) ")
CALL write_bullet_and_variable_in_CASE_NOTE ("Requesting", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE ("Pended on", pended_date)
CALL write_bullet_and_variable_in_CASE_NOTE ("Application assigned to", worker_number)
IF transfer_case = TRUE THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to team " & team_number & " in MAXIS")
CALL write_bullet_and_variable_in_CASE_NOTE ("Notes", entered_notes)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

'----------------------------------------------------------------------------------------------------EXPEDITED SCREENING!
	IF SNAP_checkbox = CHECKED THEN
    	BeginDialog exp_screening_dialog, 0, 0, 181, 165, "Expedited Screening"
      	EditBox 100, 5, 50, 15, MAXIS_case_number
      	EditBox 100, 25, 50, 15, income
      	EditBox 100, 45, 50, 15, assets
      	EditBox 100, 65, 50, 15, rent
      	CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
      	CheckBox 75, 95, 45, 10, "Electricity", electric_check
      	CheckBox 130, 95, 35, 10, "Phone", phone_check
      	ButtonGroup ButtonPressed
        	OkButton 70, 115, 50, 15
        	CancelButton 125, 115, 50, 15
      	Text 10, 140, 160, 15, "The income, assets and shelter costs fields will default to $0 if left blank. "
      	Text 5, 30, 95, 10, "Income received this month:"
      	Text 5, 50, 95, 10, "Cash, checking, or savings: "
      	Text 5, 70, 90, 10, "AMT paid for rent/mortgage:"
      	GroupBox 5, 85, 170, 25, "Utilities claimed (check below):"
      	Text 50, 10, 50, 10, "Case number: "
      	GroupBox 0, 130, 175, 30, "**IMPORTANT**"
    	EndDialog

		'DATE BASED LOGIC FOR UTILITY AMOUNTS------------------------------------------------------------------------------------------
		If date >= cdate("10/01/2017") then			'these variables need to change every October
			heat_AC_amt = 556
			electric_amt = 172
			phone_amt = 41
		Else
			heat_AC_amt = 532
			electric_amt = 141
			phone_amt = 38
		End if

    	'----------------------------------------------------------------------------------------------------THE SCRIPT
    	CALL MAXIS_case_number_finder(MAXIS_case_number)
    	Do
        	Do
    			err_msg = ""
        		Dialog exp_screening_dialog
        		cancel_confirmation
        		If isnumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbnewline & "* You must enter a valid case number."
    			If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) THEN err_msg = err_msg & vbnewline & "* The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
    			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        	LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	Loop until are_we_passworded_out = false					'loops until user passwords back in

    	''----------------------------------------------------------------------------------------------------LOGIC AND CALCULATIONS
    	'Logic for figuring out utils. The highest priority for the if...THEN is heat/AC, followed by electric and phone, followed by phone and electric separately.
    	IF heat_AC_check = checked THEN
       	utilities = heat_AC_amt
    	ELSEIF electric_check = checked and phone_check = checked THEN
       	utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
    	ELSEIF phone_check = checked and electric_check = unchecked THEN
       	utilities = phone_amt
    	ELSEIF electric_check = checked and phone_check = unchecked THEN
       	utilities = electric_amt
    	END IF

    	'in case no options are clicked, utilities are set to zero.
    	IF phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked THEN utilities = 0
    	'If nothing is written for income/assets/rent info, we set to zero.
    	IF income = "" THEN income = 0
    	IF assets = "" THEN assets = 0
    	IF rent = "" THEN rent = 0

    	'Calculates expedited status based on above numbers
    	IF (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) THEN expedited_status = "Client Appears Expedited"
    	IF (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) THEN expedited_status = "Client Does Not Appear Expedited"
    	'----------------------------------------------------------------------------------------------------checking DISQ

    	CALL navigate_to_MAXIS_screen("STAT", "DISQ")
    	'grabbing footer month and year
    	CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

    	'Reads the DISQ info for the case note.
    	EMReadScreen DISQ_member_check, 34, 24, 2
    	IF DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" THEN
       	has_DISQ = False
    	ELSE
       	has_DISQ = True
    	END IF
    	'
    	'Reads MONY/DISB to see if EBT account is open
    	IF expedited_status = "Client Appears Expedited" THEN
       	CALL navigate_to_MAXIS_screen("MONY", "DISB")
       	EMReadScreen EBT_account_status, 1, 14, 27
	   	MsgBox "This Client Appears EXPEDITED. A same day interview needs to be offered."
	   	Send_email = true
    	END IF

		IF expedited_status = "Client does not appear expedited" THEN MsgBox "This client does NOT appear expedited. A same day interview does not need to be offered."

    	'-----------------------------------------------------------------------------------------------EXPCASENOTE
    	start_a_blank_CASE_NOTE
    	CALL write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
    	CALL write_variable_in_CASE_NOTE("---")
    	CALL write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
    	CALL write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
    	CALL write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
    	CALL write_variable_in_CASE_NOTE("        Utilities (AMT/HEST claimed): $" & utilities)
    	CALL write_variable_in_CASE_NOTE("---")
    	IF has_DISQ = TRUE THEN CALL write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
    	IF has_DISQ = FALSE THEN CALL write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
    	IF expedited_status = "Client appears expedited" AND EBT_account_status = "Y" THEN CALL write_variable_in_CASE_NOTE("* EBT Account IS open.  Recipient will NOT be able to get a replacement card in the agency.  Rapid Electronic Issuance (REI) with caution.")
    	IF expedited_status = "Client appears expedited" AND EBT_account_status = "N" THEN CALL write_variable_in_CASE_NOTE("* EBT Account is NOT open.  Recipient is able to get initial card in the agency.  Rapid Electronic Issuance (REI) can be used, but only to avoid an emergency issuance or to meet EXP criteria.")
    	CALL write_variable_in_CASE_NOTE("---")
    	IF expedited_status = "Client does not appear expedited" THEN CALL write_variable_in_CASE_NOTE("Client does not appear expedited. Application sent to ECF.")
    	IF expedited_status = "Client appears expedited" THEN CALL write_variable_in_CASE_NOTE("Client appears expedited. Application sent to ECF. Emailed Triagers.")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
	END IF
'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
IF Active_checkbox = CHECKED THEN
	transfer_case = False
    action_completed = TRUE     'This is to decide if the case was successfully transferred or not
ELSE
	transfer_case = True
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMWriteScreen "X127" & worker_number, 18, 61
	transmit
	EMReadScreen worker_check, 9, 24, 2

	IF worker_check = "SERVICING" THEN
        action_completed = False
		PF10
	END IF

    EMReadScreen transfer_confirmation, 16, 24, 2
    if transfer_confirmation = "CASE XFER'D FROM" then
    	action_completed = True
    Else
        action_completed = False
    End if
END IF

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
IF send_email = True THEN CALL create_outlook_email("HSPH.EWS.Triagers@hennepin.us", "", MAXIS_case_name & maxis_case_number & " Expedited case to be assigned, transferred to team. " & worker_number & "  EOM.", "", "", TRUE)

'----------------------------------------------------------------------------------------------------NOTICE APPT LETTER Dialog
IF send_appt_ltr = TRUE THEN
    BeginDialog Hennepin_appt_dialog, 0, 0, 296, 75, "APPOINTMENT LETTER"
      EditBox 205, 25, 55, 15, interview_date
      EditBox 65, 50, 115, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 185, 50, 50, 15
        CancelButton 240, 50, 50, 15
      EditBox 65, 25, 55, 15, application_date
      Text 5, 55, 60, 10, "Worker signature:"
      Text 140, 30, 60, 10, "Appointment date:"
      GroupBox 20, 10, 255, 35, "Enter a new appointment date only if it's a date county offices are not open."
      Text 30, 30, 35, 10, "CAF date:"
    EndDialog

	'grabs CAF date, turns CAF date into string for variable
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)
	application_date = application_date & ""

	'creates interview date for 7 calendar days from the CAF date
	interview_date = dateadd("d", 7, application_date)
	If interview_date <= date then interview_date = dateadd("d", 7, date)
	interview_date = interview_date & ""		'turns interview date into string for variable
 'need to handle for if we dont need an appt letter, which would be...'

 last_contact_day = CAF_date + 30
 If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

	Do
		Do
    		err_msg = ""
    		dialog Hennepin_appt_dialog
    		cancel_confirmation
			If isdate(application_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid application date."
    		If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid interview date."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	Loop until err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

	'Figuring out the last contact day
	If app_type = "Addendum" then
	    next_month = datepart("m", dateadd("m", 1, interview_date))
	    next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
	    last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
	ELSE
	 	last_contact_day = CAF_date + 30
		If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date
    END IF

    'This checks to make sure the case is not in background and is in the correct footer month for PND1 cases.
    Do
    	call navigate_to_MAXIS_screen("STAT", "SUMM")
    	EMReadScreen month_check, 11, 24, 56 'checking for the error message when PND1 cases are not in APPL month
    	IF left(month_check, 5) = "CASES" THEN 'this means the case can't get into stat in current month
    		EMWriteScreen mid(month_check, 7, 2), 20, 43 'writing the correct footer month (taken from the error message)
    		EMWriteScreen mid(month_check, 10, 2), 20, 46 'writing footer year
    		EMWriteScreen "STAT", 16, 43
    		EMWriteScreen "SUMM", 21, 70
    		transmit 'This transmit should take us to STAT / SUMM now
    	END IF
    	'This section makes sure the case isn't locked by background, if it is it will loop and try again
    	EMReadScreen SELF_check, 4, 2, 50
    	If SELF_check = "SELF" then
    		PF3
    		Pause 2
    	End if
    Loop until SELF_check <> "SELF"

		'Navigating to SPEC/MEMO
    'Transmits to start the memo writing process

		Call start_a_new_spec_memo		'Writes the appt letter into the MEMO.
		  Call write_variable_in_SPEC_MEMO("************************************************************")
		  Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & application_date & ".")
		  Call write_variable_in_SPEC_MEMO("You need to complete an interview as part of your application.")
		  Call write_variable_in_SPEC_MEMO("The interview must be completed by " & interview_date & ".")
		  Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
		  Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & last_contact_day & " your application will be denied.") 'add 30 days
		  Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
		  Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")
		  Call write_variable_in_SPEC_MEMO("************************************************************")
		  PF4
    Call start_a_blank_CASE_NOTE
      Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO ~")
      Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of missed interview.")
      Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
      Call write_variable_in_CASE_NOTE("* Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
      'Call write_variable_in_CASE_NOTE("* A notice has been sent to client with detail about how to call in for an interview.")
      Call write_variable_in_CASE_NOTE("---")
      Call write_variable_in_CASE_NOTE(worker_signature & " via on demand waiver script")
END IF
PF3

IF action_completed = False then
    script_end_procedure ("Warning! Case did not transfer. Transfer the case manually. Script was able to complete all other steps.")
Else
    script_end_procedure ("Case has been updated please review to ensure it was processed correctly.")
End if
