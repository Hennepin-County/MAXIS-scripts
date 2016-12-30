'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EXPEDITED DETERMINATION.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
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
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS--------------------------------------------------------------------------------------------------------------------

BeginDialog Case_Number_Dialog, 0, 0, 171, 85, "Case Information"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 25, 15, elig_month
  EditBox 115, 25, 25, 15, elig_year
  EditBox 80, 45, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 60, 65, 50, 15
    CancelButton 115, 65, 50, 15
  Text 25, 10, 50, 10, "Case Number:"
  Text 5, 50, 70, 10, "Sign your case note:"
  Text 20, 30, 55, 10, "Elig month/year:"
EndDialog

BeginDialog timeliness_dialog, 0, 0, 196, 120, "Expedited Timeliness"
  EditBox 80, 5, 110, 15, date_of_application
  EditBox 80, 25, 110, 15, interview_date
  EditBox 80, 45, 110, 15, approval_date
  EditBox 10, 80, 180, 15, delay_explanation
  ButtonGroup ButtonPressed
    OkButton 85, 100, 50, 15
    CancelButton 140, 100, 50, 15
  Text 10, 50, 60, 10, "Date of Approval"
  Text 10, 10, 65, 10, "Date of Application"
  Text 10, 65, 85, 10, "Explain any delays here"
  Text 10, 30, 50, 10, "Interview Date"
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
'connecting to MAXIS & searches for the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

'dialog to gather the Case Number and such
Do
	Do
		Dialog Case_Number_Dialog
		cancel_confirmation
		err_msg = ""
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your worker signature"
		IF MAXIS_case_number = "" THEN err_msg = err_msg & vbCr & "Please enter the case number"
		IF err_msg <> "" THEN MsgBox err_msg & vbCr & vbCr & "Please resolve this to continue"
	Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false				'loops until user passwords back in

'Script is going to find information that was writen in an Expedited Screening case note using scripts
navigate_to_MAXIS_screen "CASE", "NOTE"

row = 1
col = 1
EMSearch "Received", row, col
IF row <> 0 THEN
	XFS_Screening_CNote = TRUE
	For look_for_right_note = 57 to 72
		EMReadScreen xfs_screen_note, 8, row, look_for_right_note
		IF xfs_screen_note = "expedite" THEN
			XFS_Screening_CNote = TRUE	'IF the script found a case note with the NOTES - Expedited Screening format - it can find the information used
			IF look_for_right_note = 57 or look_for_right_note = 65 THEN
				EMReadScreen xfs_screening, 32, row, 42
			ElseIf look_for_right_note = 64 OR look_for_right_note = 72 THEN
				EMReadScreen xfs_screening, 31, row, 49
			End If
			EMWriteScreen "x", row, 3
			transmit
			Exit For
		END If
	Next
ELSEIF row = 0 THEN
	XFS_Screening_CNote = FALSE
END IF

'Script is gathering the income/asset/expense information from the XFS Screening note
IF XFS_Screening_CNote = TRUE THEN
	xfs_screening = UCase(xfs_screening)
	xfs_screening_display = xfs_screening & ""

	row = 1
	col = 1
	EMSearch "CAF 1 income", row, col
	EMReadScreen caf_one_income, 8, row, 42
	IF IsNumeric(caf_one_income) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_income = abs(caf_one_income)
	Else
		caf_one_income = 0
	End If

	row = 1
	col = 1
	EMSearch "CAF 1 liquid assets", row, col
	EMReadScreen caf_one_assets, 8, row, 42
	If IsNumeric(caf_one_assets)= True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_assets = caf_one_assets * 1
	Else
		caf_one_assets = 0
	End If

	caf_one_resources = caf_one_income + caf_one_assets	'Totaling the amounts for the case note

	row = 1
	col = 1
	EMSearch "CAF 1 rent", row, col
	EMReadScreen caf_one_rent, 8, row, 42
	IF IsNumeric(caf_one_rent) = True Then 		'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_rent = abs(caf_one_rent)
	Else
		caf_one_rent = 0
	End If

	row = 1
	col = 1
	EMSearch "Utilities (amt", row, col
	EMReadScreen caf_one_utilities, 8, row, 42
	If IsNumeric(caf_one_utilities) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_utilities = abs(caf_one_utilities)
	Else
		caf_one_utilities = 0
	End If

	caf_one_expenses = caf_one_rent + caf_one_utilities		'Totaling the amounts for a case note

	'The script not adjusts the format so it looks nice
	caf_one_income = FormatCurrency(caf_one_income)
	caf_one_assets = FormatCurrency(caf_one_assets)
	caf_one_rent = FormatCurrency(caf_one_rent)
	caf_one_utilities = FormatCurrency(caf_one_utilities)
	caf_one_resources = FormatCurrency(caf_one_resources)
	caf_one_expenses = FormatCurrency(caf_one_expenses)
	PF3
End IF

'Script now goes to ELIG to find what the income/expesnse that are being used are to autofill the dialog
navigate_to_MAXIS_screen "ELIG", "FS"
EMReadScreen elig_screen_check, 4, 3, 48
IF elig_screen_check = "FSPR" Then
	transmit
	EMReadScreen is_elig_xfs, 17, 4, 3
	IF is_elig_xfs = "EXPEDITED SERVICE" THEN 	'Determines if MAXIS thinks the case is Expedited
		is_elig_xfs = TRUE
	ELSE
		is_elig_xfs = FALSE
	END IF
	is_elig_XFS = is_elig_XFS & ""
	'MsgBox is_elig_XFS

	transmit		'Finding Income and formating it
	EMReadScreen elig_gross_income, 9, 7, 72
	elig_gross_income = trim(elig_gross_income)
	elig_gross_income = abs(elig_gross_income)
	transmit

	'Finding the shelter and utility expenses and combining them and formating them
	EMReadScreen elig_heat, 3, 9, 31
	IF elig_heat = "   " THEN elig_heat = 0
	elig_heat = trim(elig_heat)
	elig_heat = abs(elig_heat)

	EMReadScreen elig_electric, 3, 8, 31
	IF elig_electric = "   " THEN elig_electric = 0
	elig_electric = trim(elig_electric)
	elig_electric = abs(elig_electric)

	EMReadScreen elig_phone, 2, 11, 32
	IF elig_phone = "  " THEN elig_phone = 0
	elig_phone = trim(elig_phone)
	elig_phone = abs(elig_phone)

	EMReadScreen elig_rent, 5, 5, 29
	IF elig_rent = "     " THEN elig_rent = 0
	elig_rent = trim(elig_rent)
	elig_rent = abs(elig_rent)

	EMReadScreen elig_tax, 5, 6, 29
	IF elig_tax = "     " THEN elig_tax = 0
	elig_tax = trim(elig_tax)
	elig_tax = abs(elig_tax)

	EMReadScreen elig_ins, 5, 7, 29
	IF elig_ins = "     " THEN elig_ins = 0
	elig_ins = trim(elig_ins)
	elig_ins = abs(elig_ins)

	EMReadScreen elig_other_exp, 5, 12, 29
	IF elig_other_exp = "     " THEN elig_other_exp = 0
	elig_other_exp = trim(elig_other_exp)
	elig_other_exp = abs(elig_other_exp)

	IF elig_heat <> 0 THEN
		elig_util = elig_heat
	ELSE
		elig_util = elig_electric + elig_phone
	END IF

	elig_shel = elig_rent + elig_tax + elig_ins + elig_other_exp
Else 		'Warning for the worker that sutofill will not work
	MsgBox ("No verision of SNAP Elig could be found for the specified month" & vbCr & "The script will not be able to gather information for you." & vbCr & vbCr & "You will need to manually enter the information for case noting")
End If

'Going to STAT for asset information
navigate_to_MAXIS_screen "STAT", "PNLR"
For pnlr_row = 3 to 19
	EMReadScreen asset_panel_type, 4, pnlr_row, 5
	IF asset_panel_type = "CASH" THEN
		EMReadScreen asset_listed, 6, pnlr_row, 26
	ELSEIF asset_panel_type = "ACCT" THEN
		EMReadScreen asset_listed, 6, pnlr_row, 31
	Else
		asset_listed = 0
	End If
	asset_amount = asset_amount + abs(trim(asset_listed))
Next

'Notifying the worker as to if the script found the XFS screening information
If XFS_Screening_CNote = FALSE Then
	MsgBox "The script could not find the Expedited Screening information in Case Notes so this information will not be included"
ElseIf XFS_Screening_CNote = True Then
	MsgBox "The script gathered infromation about the Expedited Screening, if your determination differs be sure to explain the disparity."
End If

'Prepping variables to fill in the edit boxes
determined_income = elig_gross_income & ""
determined_assets = asset_amount & ""
determined_shel = elig_shel & ""
determined_utilities = elig_util & ""

'THIS DIALOG IS DEFINED HERE BECAUSE OTHERWISE THE SCRIPT DOES NOT AUTOFILL THE TEXT FIELDS THAT ARE VAIRABLES
BeginDialog Expedited_Detail_Dialog, 0, 0, 401, 325, "Expedited Determination"
  EditBox 265, 50, 120, 15, determined_income
  EditBox 240, 70, 145, 15, determined_assets
  EditBox 275, 90, 110, 15, determined_shel
  EditBox 275, 110, 110, 15, determined_utilities
  DropListBox 300, 130, 85, 20, "TRUE"+chr(9)+"FALSE", is_elig_XFS
  CheckBox 200, 150, 150, 10, "Check here if APPLICANT has no form of ID", id_check
  EditBox 15, 180, 370, 15, out_of_state_explanation
  EditBox 15, 215, 370, 15, previous_xfs_explanation
  EditBox 15, 250, 370, 15, abawd_explanation
  EditBox 15, 285, 370, 15, other_explanation
  ButtonGroup ButtonPressed
    OkButton 290, 305, 50, 15
    CancelButton 345, 305, 50, 15
  Text 205, 55, 50, 10, "Gross Income:"
  Text 205, 75, 25, 10, "Assets:"
  Text 205, 95, 60, 10, "Shelter Expense:"
  Text 205, 115, 60, 10, "Utilities Expense:"
  Text 200, 135, 85, 10, "Client appears Expedited:"
  Text 10, 95, 160, 10, xfs_screening
  GroupBox 195, 5, 195, 140, "Information from SNAP/ELIG"
  GroupBox 5, 5, 180, 105, "Expedited Screening"
  Text 10, 20, 145, 10, "Information pulled from previous case note."
  Text 15, 50, 60, 10, "Assets from CAF1:"
  Text 15, 65, 90, 10, "Rent/Mortgage from CAF1:"
  Text 15, 80, 65, 10, "Utilities from CAF1:"
  Text 110, 35, 80, 10, caf_one_income
  Text 110, 50, 75, 10, caf_one_assets
  Text 110, 65, 65, 10, caf_one_rent
  Text 200, 35, 180, 10, "This information can be altered for the case note."
  Text 15, 35, 65, 10, "Income from CAF1: "
  Text 10, 115, 170, 35, "If the Expedited Determination for screening and elig do not match, detail the information that changed the determination from what is on CAF1 to the final determination."
  Text 10, 270, 95, 10, "Other detail needed to clarify"
  Text 110, 80, 75, 10, caf_one_utilities
  Text 10, 165, 255, 10, "If client received SNAP benefits out of state that impact eligibility, explain here"
  Text 200, 20, 180, 10, "Information based on current STAT and ELIG panels"
  Text 10, 200, 395, 10, "If the last issuance client received was Expedited and delayed verifications were not provided, explain in detail here:"
  Text 10, 235, 140, 10, "If client is an ABAWD, provide detail here:"
EndDialog

'Running the Dialog asking for all the detail and explanations
DO
	Do
		Dialog Expedited_Detail_Dialog
		cancel_confirmation
		err_msg = ""
		IF is_elig_XFS = "FALSE" AND out_of_state_explanation = "" AND previous_xfs_explanation = "" AND other_explanation = "" AND abawd_explanation = "" THEN err_msg = err_msg & vbCr & "You have determined this case to NOT be Expedited but have provided no detail explanation" & vbCr & "Please complete at least one of the explanation boxes."
		IF id_check = checked AND other_explanation = "" THEN err_msg = err_msg & vbCr & "Please provided detail about no ID, remember that this is ONLY for the applicant and does NOT need to be a photo ID"
		IF err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Formating the information from the edit boxes
If determined_income = "" Then determined_income = 0
If determined_assets = "" Then determined_assets = 0
If determined_shel = "" Then determined_shel = 0
If determined_utilities = "" Then determined_utilities = 0
determined_resources = abs(determined_income) + (determined_assets * 1)
determined_expenses = abs(determined_shel) + abs(determined_utilities)
determined_assets = FormatCurrency(determined_assets)
determined_expenses = FormatCurrency(determined_expenses)
determined_income = FormatCurrency(determined_income)
determined_resources = FormatCurrency(determined_resources)
determined_shel = FormatCurrency(determined_shel)
determined_utilities = FormatCurrency(determined_utilities)

'Converting String entries to Boolean
IF is_elig_XFS = "TRUE" Then is_elig_XFS = TRUE
IF is_elig_XFS = "FALSE" Then is_elig_XFS = FALSE

'Dialog about timeliness will run if case is determined to be expedited
IF is_elig_XFS = TRUE Then
	Do
		Do
			Do
				Dialog timeliness_dialog
				cancel_confirmation
				err_msg = ""
				IF date_of_application = "" OR IsDate(date_of_application) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid date of application."
				IF interview_date = "" OR IsDate(interview_date) = FALSE Then err_msg = err_msg & vbCr & "Pleaes enter a valid Interview Date."
				IF approval_date = "" OR IsDate(approval_date) = FALSE Then err_msg = err_msg & vbCr & "Please enter a valid Date of Approval."
				IF err_msg <> "" Then MsgBox err_msg
			Loop until err_msg = ""
			days_delayed = DateDiff ("d", date_of_application, approval_date)
			IF days_delayed > 7 AND delay_explanation = "" Then err_msg = err_msg & vbCr & "Your approval is more than 7 days from the date of application." & vbCr & "Please provide an explanation for the delay."
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
End If

'creating a custom header: this is read by BULK - EXP SNAP REVIEW script so don't mess this please :)
IF is_elig_XFS = true then
	case_note_header_text = "Expedited Determination: SNAP appears expedited"
ELSEIF is_elig_XFS = False then
	case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
END IF

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------
navigate_to_MAXIS_screen "CASE", "NOTE"
Call start_a_blank_case_note
Call write_variable_in_case_note (case_note_header_text)
IF XFS_Screening_CNote = TRUE Then
	Call write_bullet_and_variable_in_case_note ("Expedited Screening found", xfs_screening)
	Call write_variable_in_case_note ("*   Based on: Income: " & caf_one_income & ",   Assets: " & caf_one_assets & ",  Totaling: " & caf_one_resources)
	Call write_variable_in_case_note ("*             Shelter: " & caf_one_rent & ", Utilities: " & caf_one_utilities & ", Totaling: " & caf_one_expenses)
	Call write_variable_in_case_note ("---")
End If
If interview_date <> "" Then Call write_variable_in_case_note ("* Interview completed on: " & interview_date & " and full Expedited Determination Done")
IF is_elig_XFS = TRUE Then Call write_variable_in_case_note ("* Case is determined to meet criteria and Expedited SNAP can be approved.")
IF is_elig_XFS = FALSE Then Call write_variable_in_case_note ("* Expedited SNAP cannot be approved as case does not meet all criteria")
Call write_variable_in_case_note ("*   Based on: Income: " & determined_income & ",   Assets: " & determined_assets & ",   Totaling: " & determined_resources)
Call write_variable_in_case_note ("*             Shelter: " & determined_shel & ", Utilities: " & determined_utilities & ",  Totaling: " & determined_expenses)
IF id_check = checked Then Call write_variable_in_case_note ("* Applicant has not provided proof of ID.")
IF out_of_state_explanation <> "" Then
	Call write_variable_in_case_note ("* SNAP benefits have been received in another state")
	Call write_variable_in_case_note ("*    " & out_of_state_explanation)
End If
If previous_xfs_explanation <> "" Then
	Call write_variable_in_case_note ("* Expedited SNAP was the last approval and delayed verifs were not received")
	Call write_variable_in_case_note ("*    " & previous_xfs_explanation)
End If
Call write_bullet_and_variable_in_case_note("ABAWD info/explanation", abawd_explanation)
Call write_bullet_and_variable_in_case_note ("Other Notes", other_explanation)
Call write_variable_in_case_note ("---")
IF is_elig_XFS = TRUE Then
	Call write_bullet_and_variable_in_case_note ("Date of Application", date_of_application)
	Call write_bullet_and_variable_in_case_note ("Date of Interview", interview_date)
	Call write_bullet_and_variable_in_case_note ("Date of Approval", approval_date)
	Call write_bullet_and_variable_in_case_note ("Reason for Delay", delay_explanation)
	Call write_variable_in_case_note ("---")
End If
Call write_variable_in_case_note(worker_signature)

script_end_procedure ("Success! The script is complete. Case note has been entered detailing your Expedited Determination.")
