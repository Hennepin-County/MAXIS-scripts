'Required for statistical purposes===============================================================================
name_of_script = "BULK - INAC SCRUBBER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 169                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================


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

'THE FOLLOWING VARIABLE IS DYNAMICALLY DETERMINED BY THE PRESENCE OF DATA IN CLS_x1_number. IT WILL BE ADDED DYNAMICALLY TO THE DIALOG BELOW.
If CLS_x1_number <> "" then CLS_dialog_string = "**This script will XFER cases in REPT/INAC to " & CLS_x1_number & ".**"

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

'Shows the dialog, requires 7 digits for worker number, a worker signature, a MAXIS_footer_month and year. Contains developer mode to bypass case noting and XFERing.
Do
	'Adding err_msg handling
	err_msg = ""

	'DIALOG----------------------------------------------------------------------------------------------------
	'NOTE: this dialog uses a dynamic CLS_dialog_string variable. As such, it can't be directly edited in dialog editor.
	'Adding Dialog1 for infinite dlg potential (possible monthly workflow, yo)
	BeginDialog Dialog1, 0, 0, 206, 162, "INAC Scrubber Dialog"
	  EditBox 80, 80, 80, 15, worker_signature
	  EditBox 145, 100, 60, 15, worker_number
	  EditBox 55, 120, 35, 15, MAXIS_footer_month
	  EditBox 145, 120, 35, 15, MAXIS_footer_year
	  ButtonGroup ButtonPressed
	  	OkButton 45, 140, 50, 15
	  	CancelButton 110, 140, 50, 15
	  Text 5, 5, 200, 10, CLS_dialog_string
	  Text 5, 25, 195, 20, "Script will check MMIS for each household memb, ABPS for Good Cause status, and CCOL/CLIC for claims."
	  Text 5, 55, 195, 20, "Write the information in the boxes below and click ''OK'' to begin. Click ''Cancel'' to exit."
	  Text 5, 85, 75, 10, "Sign your case notes:"
	  Text 5, 105, 135, 10, "Write your worker number (7 digit format):"
	  Text 5, 125, 45, 10, "Footer month:"
	  Text 100, 125, 40, 10, "Footer year:"
	EndDialog

	Dialog
	cancel_confirmation
	If worker_signature = "UUDDLRLRBA" then
		MsgBox "Developer mode enabled. Will bypass XFER and case note functions."
		worker_signature = ""
		developer_mode = True
	Else
		developer_mode = False
	End if
	If len(worker_number) <> 7 								THEN err_msg = err_msg & vbNewLine & "* Your worker number is not 7 digits. Please try again. Type the whole worker number."
	IF worker_signature = "" AND developer_mode = FALSE		THEN err_msg = err_msg & vbNewLine & "* You must sign your case notes. Please provide a worker signature."
	If MAXIS_footer_month = "" or MAXIS_footer_year = "" 	THEN err_msg = err_msg & vbNewLine & "* You must provide a footer month and year!"
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Converts system/footer month and year to a MAXIS-appropriate number, for validation
current_system_month = DatePart("m", Now)
If len(current_system_month) = 1 then current_system_month = "0" & current_system_month
current_system_year = DatePart("yyyy", Now) - 2000
If len(MAXIS_footer_month) <> 2 or isnumeric(MAXIS_footer_month) = False or MAXIS_footer_month > 13 or len(MAXIS_footer_year) <> 2 or isnumeric(MAXIS_footer_year) = False then script_end_procedure("Your footer month and year must be 2 digits and numeric. The script will now stop.")
footer_month_first_day = MAXIS_footer_month & "/01/" & MAXIS_footer_year
date_compare = datediff("d", footer_month_first_day, date)
If date_compare < 0 then script_end_procedure("You appear to have entered a future month and year. The script will now stop.")
IF developer_mode = FALSE THEN
	If cint(current_system_month) = cint(MAXIS_footer_month) and cint(MAXIS_footer_year) = cint(current_system_year) AND developer_mode = False then script_end_procedure("Do not use this script in the current footer month. These cases need to be in your REPT/INAC for 30 days. The script will now stop.")
END IF

'Warning message before executing
warning_message = MsgBox(	"Worker: " & worker_number & vbCr & _
							"Footer month/year: " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbCr & _
							vbCr & _
							"This script will case note EACH case on the above REPT/INAC, in the selected footer month, and XFER to " & CLS_x1_number & ", under the following conditions:" & vbCr & _
							"   " & chr(183) & " Case has no open HC on this case number. " & vbCr & _
							"   " & chr(183) & " Case has no open IMA. " & vbCr & _
							"   " & chr(183) & " Case has HC that did not close for no-or-incomplete renewal. " & vbCr & _
							"   " & chr(183) & " Case has no messages currently on the DAIL. " & vbCr & _
							"   " & chr(183) & " Case is a closure, and not a denial. For denials, use ''Denied progs''. " & vbCr & _
							vbCr & _
							"This script will also generate a Word document with the following info from the entire caseload: " & vbCr & _
							"   " & chr(183) & " CCOL/CLIC information. " & vbCr & _
							"   " & chr(183) & " Good cause ABPS status. " & vbCr & _
							"   " & chr(183) & " Privileged cases you cannot access. " & vbCr & _
							vbCr & _
							"It requires the use of MDHS for your state systems log-on, as it needs to check MMIS. Also, it only runs in the month before the current footer month (or any month prior)." & vbCr & _
							vbCr & _
							"Please press OK to continue, or cancel to exit the script.", vbOKCancel)
If warning_message = vbCancel then stopscript

'Connects to (and checks for) MAXIS
EMConnect ""
check_for_MAXIS(True)

'Gets back to SELF
back_to_self

'Reads if we're in training or production. If we're in production, it'll use "10" (regular MMIS) as an MMIS code. If we're in training, it'll use "11" (training MMIS) as an MMIS code.
EMReadScreen environment_check, 12, 22, 48
If trim(environment_check) = "TRAINING" then
	If developer_mode = False then MsgBox "Training mode detected. Will check MMIS/case info in training region."	'Only alerts non-developers
	MMIS_number = "11"
	MMIS_MDHS_row = 16
	MAXIS_number = "3"
	MAXIS_MDHS_row = 8
ElseIf trim(environment_check) = "PRODUCTION" then
	MMIS_number = "10"
	MMIS_MDHS_row = 15
	MAXIS_number = "1"
	MAXIS_MDHS_row = 6
ElseIf trim(environment_check) = "INQUIRY DB" and developer_mode = True then		'Can use inquiry if you're on developer mode.
	MMIS_number = "10"
	MMIS_MDHS_row = 15
	MAXIS_number = "2"
	MAXIS_MDHS_row = 7
Else
	script_end_procedure("This script must be run on either MAXIS training or MAXIS production. Please try again.")
End if

'Enters REPT/INAC under the specific footer month and year, clearing any case number currently loaded.
EMWriteScreen "REPT", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
EMWriteScreen "INAC", 21, 70
transmit

'Checks to make sure the selected worker number is the default. If not it will navigate to that person.
EMReadScreen worker_number_check, 7, 21, 16
If ucase(worker_number_check) <> ucase(worker_number) then
	EMWriteScreen worker_number, 21, 16
	transmit
End if

'Checks to make sure the worker has cases to close. If not the script will end.
EMReadScreen worker_has_cases_to_close_check, 16, 7, 14
If worker_has_cases_to_close_check = "                " then script_end_procedure("This worker does not appear to have any cases to close. If there are cases here email the script alpha user a description of the problem and your worker number.")

'Notifies the worker that we're about to create a Word document.
word_warning = MsgBox("The script is about to start a Word document. This may take a few moments.", vbOKCancel)
If word_warning = vbCancel then stopscript

'Before creating the Word document, we create an Excel spreadsheet that runs behind the scenes to collect the case numbers. It's easier to work off of an Excel spreadsheet than an array (for debugging purposes). An array would be faster, however.

'Loads Excel document for developer_mode only
If developer_mode = True then
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True

	'Assigning values to the Excel spreadsheet.
	ObjExcel.Cells(1, 1).Value = "MAXIS number"
	ObjExcel.Cells(1, 2).Value = "Name"
	ObjExcel.Cells(1, 3).Value = "INAC eff date"
	ObjExcel.Cells(1, 4).Value = "Amount due in claims"
	ObjExcel.Cells(1, 5).Value = "DAILs?"
	ObjExcel.Cells(1, 6).Value = "MMIS?"
	ObjExcel.Cells(1, 7).Value = "PMIs"
	ObjExcel.Cells(1, 8).Value = "Privileged?"
	ObjExcel.Cells(1, 9).Value = "HC renewal closure?"
	objExcel.Cells(1, 10).Value = "Transfer Case?"
End if

'Now it creates a Word document to store all of the active claims.
Set objWord = CreateObject("Word.Application")
objWord.Visible = true
set objDoc = objWord.Documents.add()
Set objSelection = objWord.Selection
objselection.typetext "Case numbers with active claims: "
objselection.TypeParagraph()
objselection.TypeParagraph()


'Setting the variable for the do...loop.
MAXIS_row = 7 'This sets the variable for the following do...loop.

'This loop grabs the case number, client name, and inac date for each case.
Do
	Do
		EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3          'First it reads the case number, name, date they closed, and the APPL date.
		EMReadScreen client_name, 25, MAXIS_row, 14
		EMReadScreen inac_date, 8, MAXIS_row, 49
		EMReadScreen appl_date, 8, MAXIS_row, 39
		MAXIS_case_number = Trim(MAXIS_case_number)                    'Then it trims the spaces from the edges of each.
		client_name = Trim(client_name)
		inac_date = Trim(inac_date)
		appl_date = Trim(appl_date)
		If appl_date <> inac_date then                     'Because if the two dates equal each other, then this is a denial and not a case closure.

			'Adds case info to an array. Uses tildes to differentiate the MAXIS_case_number, client_name, and inac_date. Uses vert lines to differentiate entries. Will be fleshed out later.
			If INAC_info_array = "" then
				INAC_info_array = MAXIS_case_number & "~" & client_name & "~" & inac_date
			Else
				INAC_info_array = INAC_info_array & "|" & MAXIS_case_number & "~" & client_name & "~" & inac_date
			End if
		End if
		MAXIS_row = MAXIS_row + 1
	Loop until MAXIS_row = 19
	MAXIS_row = 7 'Setting the variable for when the do...loop restarts
	PF8
	EMReadScreen last_page_check, 21, 24, 2 'checks for "THIS IS THE LAST PAGE"
Loop until last_page_check = "THIS IS THE LAST PAGE"

'Splits INAC_info_array into individual case numbers
INAC_info_array = split(INAC_info_array, "|")

'Creates a variable for total_cases, which will be used by the various for...nexts
total_cases = ubound(INAC_info_array)

'Declares INAC_scrubber_primary_array, redims it to be the size needed for our total amount of cases.
Dim INAC_scrubber_primary_array()
ReDim INAC_scrubber_primary_array(total_cases, 8)
' 0 = MAXIS case #
' 1 = CL Name
' 2 = INAC Date
' 3 = Claims?
' 4 = DAILs?
' 5 = MMIS Status
' 6 = PMI array
' 7 = privileged
' 8 = to transfer or not to transfer b/c of HC...that is the question...true means TRANSFER, false means NO TRANSFER


'Assigns info to the array. If developer_mode is on, it'll also add to an Excel spreadsheet
For x = 0 to total_cases
	interim_array = split(INAC_info_array(x), "~")			'This is a temporary array, and is always three objects (MAXIS_case_number, client_name, INAC_date)
	INAC_scrubber_primary_array(x, 0) = interim_array(0)	'The MAXIS_case_number
	INAC_scrubber_primary_array(x, 1) = interim_array(1)	'The client_name
	INAC_scrubber_primary_array(x, 2) = interim_array(2)	'The inac_date
	If developer_mode = True then
		ObjExcel.Cells(x + 2, 1).Value = INAC_scrubber_primary_array(x, 0)
		ObjExcel.Cells(x + 2, 2).Value = INAC_scrubber_primary_array(x, 1)
		ObjExcel.Cells(x + 2, 3).Value = INAC_scrubber_primary_array(x, 2)
	End if
	'Setting a default value for (x, 8)
	INAC_scrubber_primary_array(x, 8) = TRUE
Next

'Navigates to CCOL/CLIC
EMWriteScreen "CCOL", 20, 22
EMWriteScreen "CLIC", 20, 70
transmit

'Grabs any claims due for each case. Adds to Excel if developer_mode = True
For x = 0 to total_cases
	MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
	EMWriteScreen "________", 4, 8
	EMWriteScreen MAXIS_case_number, 4, 8
	transmit
	EMReadScreen claims_due, 10, 19, 58
	INAC_scrubber_primary_array(x, 3) = claims_due
	If developer_mode = True then ObjExcel.Cells(x + 2, 4).Value = claims_due
Next

'Entering claims into the Word doc
For x = 0 to total_cases
	'Grabbing the MAXIS_case_number, client_name, and claims_due from the array
	MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
	client_name = INAC_scrubber_primary_array(x, 1)
	claims_due = INAC_scrubber_primary_array(x, 3)

	'If there's a claim due, it'll add to the Word doc
	If claims_due <> 0 then
		objselection.typetext MAXIS_case_number & ": " & client_name & "; amount due: $" & claims_due
		objselection.TypeParagraph()
	End if
Next

'Navigating to the DAIL (Goes back to self as there are issues in CCOL/CLIC with the direct navigate_to_MAXIS_screen)
back_to_SELF
call navigate_to_MAXIS_screen("DAIL", "DAIL")

'This checks the DAIL for messages, sends a variable to the array. We don't transfer cases with DAIL messages. (True for "has DAIL", False for "doesn't have DAIL")
For x = 0 to total_cases
	MAXIS_case_number = INAC_scrubber_primary_array(x, 0)		'Grabbing case number
	EMWriteScreen "________", 20, 38
	EMWriteScreen MAXIS_case_number, 20, 38
	transmit
	EMReadScreen DAIL_check, 1, 5, 5
	If DAIL_check <> " " then
		INAC_scrubber_primary_array(x, 4) = True
	Else
		INAC_scrubber_primary_array(x, 4) = False
	End if
	If developer_mode = True then ObjExcel.Cells(x + 2, 5).Value = INAC_scrubber_primary_array(x, 4)
	excel_row = excel_row + 1
Next

'Making the header for the next section of the Word document.
objselection.TypeParagraph()
objselection.TypeParagraph()
objselection.typetext "Cases that need to be REINed, STAT/ABPS updated with an ''N'' code for Good Cause Status, and then reapproved for closure:"
objselection.TypeParagraph()

'This do...loop goes into STAT, grabs PMIs for MEMB types 01, 02, 03, 04, and 18, and then navigates to ABPS to get that info.
For x = 0 to total_cases
	'Grabbing case number and name for this loop
	MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
	client_name = INAC_scrubber_primary_array(x, 1)

	'Gets to MEMB
	call navigate_to_MAXIS_screen("STAT", "MEMB")

	'Checks to make sure we're past SELF. If we aren't, it'll save that the case is privileged (most likely cause) in the array.
	EMReadScreen SELF_check, 4, 2, 50
	If SELF_check = "SELF" then
		INAC_scrubber_primary_array(x, 7) = True		'If it's privileged, it won't get past SELF. If so, it'll enter a True value in the array.
		objSelection.TypeText MAXIS_case_number & ": Case is privileged. Cannot transfer."
		objSelection.TypeParagraph()
	Else
		INAC_scrubber_primary_array(x, 7) = False		'If it gets through, it isn't privileged.

		'Reads the PMI number for each 01, 02, 03, 04, and 18 (dependents). Stores it to a variable called "PMI array", which will be added to the main array (and split later).
		Do
			EMReadScreen PMI_number, 8, 4, 46
			EMReadScreen rel_to_applicant, 2, 10, 42
			If rel_to_applicant = "01" or rel_to_applicant = "02" or rel_to_applicant = "03" or rel_to_applicant = "04" or rel_to_applicant = "18" then
				If PMI_array = "" then
					PMI_array = trim(PMI_number)
				Else
					PMI_array = PMI_array & "|" & trim(PMI_number)
				End if

			End if
			transmit
			EMReadScreen no_more_MEMBs_check, 31, 24, 2
			INAC_scrubber_primary_array(x, 6) = PMI_array		'Writes the PMI array to the main array
		Loop until no_more_MEMBs_check = "ENTER A VALID COMMAND OR PF-KEY"

		'Goes to ABPS to check good cause. Good cause will not hang the case from being sent to CLS, as such, it does not get entered in the array (just the Word doc).
		call navigate_to_MAXIS_screen("STAT", "ABPS")
		EMReadScreen good_cause_check, 1, 5, 47
		If good_cause_check = "P" then
			objselection.typetext MAXIS_case_number & ", " & client_name
			objselection.TypeParagraph()
		End if
	End if
	If developer_mode = True then
		ObjExcel.Cells(x + 2, 8).Value = INAC_scrubber_primary_array(x, 7)		'Writes privileged status to Excel when developer_mode is on
		ObjExcel.Cells(x + 2, 7).Value = INAC_scrubber_primary_array(x, 6)		'Writes PMI array to Excel when developer_mode is on
	End if
	PMI_array = ""		'Clears the variable for the following loop
Next

'MMIS--------------------------------------------------------------------------------------------------------------

'The following checks for which screen MMIS is running on.
'If the user does not have MMIS access, they can bypass MMIS run mode.
attn
DO
	EMReadScreen MMIS_A_check, 7, MMIS_MDHS_row, 15
	IF MMIS_A_check = "RUNNING" then
		EMSendKey MMIS_number
		transmit
		mmis_mode = TRUE
		EXIT DO
	ELSEIF MMIS_A_check <> "RUNNING" then
		attn
		'Looking for other BlueZone sessions
		session_b = EMConnect ("B")
		'If the user has an S2 running, the script will look for it and connect to it and check for MMIS...
		IF session_b = 0 THEN
			EMConnect "B"
			attn
			EMReadScreen MMIS_B_check, 7, MMIS_MDHS_row, 15
			If MMIS_B_check <> "RUNNING" then
				no_mmis = MsgBox ("MMIS does not appear to be running." & vbNewLine & _
							"  *If this is correct and you do not have MMIS access, press ''Ignore.''" & vbNewLine & _
							"  *If you need to start MMIS, please do so and THEN press ''Retry.''" & vbNewLine & _
							"  *Otherwise, press ''Abort'' to stop the script.", vbAbortRetryIgnore + vbInformation + vbSystemModal + vbDefaultButton2, "MMIS Not Found")
				IF no_mmis = vbAbort THEN script_end_procedure("Script cancelled for no MMIS.")
				IF no_mmis = vbIgnore THEN
					mmis_mode = FALSE
					EXIT DO
				END IF
			End if
			If MMIS_B_check = "RUNNING" then
				EMSendKey MMIS_number
				transmit
				mmis_mode = TRUE
				EXIT DO
			End if
		'Otherwise, if the user does not have an S2 running...
		Else
			no_mmis = MsgBox ("MMIS does not appear to be running." & vbNewLine & _
						"  *To stop the script, press ''Abort.''" & vbNewLine & _
						"  *If you need to start MMIS, please do so and THEN press ''Retry'' for the script to continue." & vbNewLine & _
						"  *Otherwise, press ''Ignore'' to have the script continue without accessing MMIS.", vbAbortRetryIgnore + vbInformation + vbSystemModal + vbDefaultButton3, "MMIS Not Found")
			IF no_mmis = vbAbort THEN script_end_procedure("Script cancelled for no MMIS.")
			IF no_mmis = vbIgnore THEN
				mmis_mode = FALSE
				EXIT DO
			END IF
		END IF
	End if
LOOP

'If the user has opted out of running the script through MMIS
IF mmis_mode = TRUE THEN
	'Shifts user focus to whatever screen ended up getting selected (A or B)
	EMFocus

	'Sending MMIS back to the beginning screen and checking for a password prompt
	Do
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then StopScript
		EMReadScreen session_start, 18, 1, 7
	Loop until session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
	EMWriteScreen "mw00", 1, 2
	transmit
	transmit

	'Finding the right MMIS, if needed, by checking the header of the screen to see if it matches the security group selector
	EMReadScreen MMIS_security_group_check, 21, 1, 35
	If MMIS_security_group_check = "MMIS MAIN MENU - MAIN" then
		EMSendKey "x"
		transmit
	End if

	'Now it finds the recipient file application feature and selects it.
	row = 1
	col = 1
	EMSearch "RECIPIENT FILE APPLICATION", row, col
	EMWriteScreen "x", row, col - 3
	transmit

	'Now we are in RKEY, and it enters an I
	EMWriteScreen "i", 2, 19

	'This for...next enters a PMI for each HH member and gets their program status in MMIS.
	For x = 0 to total_cases
		'Grabs MAXIS_case_number and PMI_array from the main array
		MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
		PMI_array = INAC_scrubber_primary_array(x, 6)
		privileged_status = INAC_scrubber_primary_array(x, 7)

		'Splits the PMI_array from the main array into an actual array, which will be used by the script.
		PMI_array = split(PMI_array, "|")

		'Checks each PMI in MMIS for discrepancies. Skips cases with no PMI (privileged cases).
		For each PMI in PMI_array
			If privileged_status = False then
				If len(PMI) < 8 then 'This will generate an 8 digit PMI.
					Do
						PMI = "0" & PMI
					Loop until len(PMI) = 8
				End if
				EMWriteScreen PMI, 4, 19
				transmit
				EMWriteScreen "RELG", 1, 8
				transmit
				EMReadScreen MMIS_case_status, 1, 7, 62
				If len(MAXIS_case_number) < 8 then 'This will generate an 8 digit MAXIS case number.
					Do
						MAXIS_case_number = "0" & MAXIS_case_number
					Loop until len(MAXIS_case_number) = 8
				End if
				EMReadScreen MMIS_case_number, 8, 6, 73

				'Checks for active/pending cases. Shows "True" for MMIS is active, and "False" for MMIS is not active.
				'It considers IMA cases (cases which start with a 25) to be connected to MAXIS (we don't want to lose these cases), and will exclude them from the MAXIS pieces.
				'If it's closed/denied (C or D), if there's no end date, it'll consider it connected to MAXIS (discrepancy handling)
				If MMIS_case_status = "A" or MMIS_case_status = "P" then
					If isnumeric(MMIS_case_number) = False or (MMIS_case_number = MAXIS_case_number) or left(MMIS_case_number, 2) = "25" then
						INAC_scrubber_primary_array(x, 5) = True
					End if
				ElseIf MMIS_case_status = "C" or MMIS_case_status = "D" then
					EMReadScreen elig_type, 2, 6, 33
					EMReadScreen elig_end_date, 8, 7, 36
					If elig_end_date = "99/99/99" then
						INAC_scrubber_primary_array(x, 5) = True
					Else		'Allows for cases that are closing next month
						If datediff("m", elig_end_date, now) < 1 and (isnumeric(MMIS_case_number) = False or MMIS_case_number = MAXIS_case_number) then
							INAC_scrubber_primary_array(x, 5) = True
						ELSE
							INAC_scrubber_primary_array(x, 5) = FALSE
						End if
					End if
				End if
				PF6
			End if
		Next
		If INAC_scrubber_primary_array(x, 5) <> True then INAC_scrubber_primary_array(x, 5) = False		'Sets this after the others, so that it doesn't refresh each loop.
		If developer_mode = True then
			FOR asdf = 0 TO 8
				objExcel.Cells(x + 2, asdf + 1).Value = INAC_scrubber_primary_array(x, asdf)		'Writes MMIS status to array when developer_mode is on
			NEXT
		END IF
	Next

	'The following checks for which screen MAXIS is running on.
	EMConnect "A"
	attn
	EMReadScreen MAXIS_A_check, 7, MAXIS_MDHS_row, 15
	IF MAXIS_A_check = "RUNNING" then
		EMSendKey MAXIS_number
		transmit
	End if
	IF MAXIS_A_check <> "RUNNING" then
		attn
		EMConnect "B"
		EMReadScreen MAXIS_B_check, 7, MAXIS_MDHS_row, 15
		If MAXIS_B_check <> "RUNNING" then
			MsgBox "The script is struggling to find MAXIS. Please navigate back to MAXIS and press OK for the script to continue."
		Else
			EMSendkey MAXIS_number
			transmit
		End if
	End if
END IF

'Header for the MMIS discrepancies section of the doc
objSelection.TypeParagraph()
objselection.typetext "Case numbers not transferred because of HC: "
objselection.TypeParagraph()
objselection.TypeParagraph()

'This do...loop updates case notes for all of the cases that don't have DAIL messages or cases still open in MMIS
For x = 0 to total_cases
	'Grabs MAXIS_case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
	MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
	DAILS_out = INAC_scrubber_primary_array(x, 4)
	MMIS_status = INAC_scrubber_primary_array(x, 5)
	privileged_status = INAC_scrubber_primary_array(x, 7)

	'Reseting this value for CASE NOTING and TIKLing reasons
	closure_reason = ""
	closure_date = ""
	inac_month = ""

	'Adds the case number to word doc if MMIS is active
	If MMIS_status = true Then
		objselection.typetext MAXIS_case_number
		objselection.TypeParagraph()
	End If

	'Checking to determine that the client closed for no or incomplete review. If that is the case, then the script does not transfer the client to CLS
	'This is also where the script will check to see that HC closed in the most recent month. If HC closed and the user bypassed MMIS, the script will skip over this case.

	CALL navigate_to_MAXIS_screen("CASE", "CURR")
	EMWriteScreen "X", 4, 9
	transmit

	'Need to check for HC closing this month.
	'First, we are going to write "MA" at 3, 19. IF MA is not found to have closed this month, then we
	' ... will write "QM" then "SL" then "Q1"...

	'Creating an array of HC programs to check.
	'We would put in a policy citation but the new manual does not have numbers because that would be too easy.
	'See EPM for details.
	hc_progs_array = "MA,QM,SL,Q1"
	hc_progs_array = split(hc_progs_array, ",")
	FOR EACH maxis_hc_program IN hc_progs_array
		CALL write_value_and_transmit(maxis_hc_program, 3, 19)

		EMReadScreen closure_reason, 9, 8, 60
		EMReadScreen closure_date, 5, 8, 28
		EMReadScreen inac_month, 5, 20, 54
		inac_month = replace(inac_month, " ", "/")

		'Checking to determine if this program closed in the INAC month.
		IF inac_month = closure_date THEN
			'if worker bypassed MMIS we are marking case to keep them.
			IF mmis_mode = FALSE THEN
				objSelection.typetext MAXIS_case_number & ": case has HC that closed this month. Please review manually as MMIS was bypassed."
				objSelection.TypeParagraph()
				INAC_scrubber_primary_array(x, 8) = FALSE
				EXIT FOR
			END IF
			'Otherwise, if the user navigates to MMIS, the script will check for "NO REVIEW" as the reason for closure.
			IF closure_reason = "NO REVIEW" THEN
				'If "NO REVIEW" is found as the reason for the closure, the script will hang on to the case
				INAC_scrubber_primary_array(x, 8) = FALSE
				objSelection.typetext MAXIS_case_number & ": case has HC client(s) that closed for incomplete or no review. Policy gives CL 4-month reinstate period."
				objSelection.TypeParagraph()
				EXIT FOR
			END IF
		END IF
	NEXT

	'Reseting values in the Excel spreadsheet
	IF developer_mode = True THEN
		FOR asdf = 0 TO 8
			objExcel.Cells(x + 2, asdf + 1).Value = INAC_scrubber_primary_array(x, asdf)
		NEXT
	END IF

	back_to_self

	MMIS_status = INAC_scrubber_primary_array(x, 5)

	'Now determining whether or not the script is going to transfer this case. This determines the verbiage in CASE NOTE
	'If we are to this point and we have not said that we are holding on to this case (b/c of HC considerations) then
	' ... we need to determine if there are other reasons why we would not transfer this case...
	IF INAC_scrubber_primary_array(x, 8) = TRUE THEN
		'If it is privileged, we cannot transfer
		IF privileged_status = TRUE THEN
			INAC_scrubber_primary_array(x, 8) = FALSE
		ELSE
			'If there are DAILS, we cannot transfer
			IF DAILS_out = TRUE THEN
				INAC_scrubber_primary_array(x, 8) = FALSE
			ELSE
				'If we have gone into MMIS and found that the case has an active HC program that we care aboot then we are hanging on to this case
				IF MMIS_status = TRUE AND mmis_mode = TRUE THEN INAC_scrubber_primary_array(x, 8) = FALSE
			END IF
		END IF
	END IF

	'The case notey gobbins...
	'If we are going to transfer this case, then we get the following case note...
	If privileged_status <> TRUE AND INAC_scrubber_primary_array(x, 8) = TRUE then
		call navigate_to_MAXIS_screen("CASE", "NOTE")
		PF9
		If developer_mode = False then
			call write_variable_in_case_note("--------------------Case is closed--------------------")
			call write_variable_in_case_note("* Reviewed closed case for claims via automated script.")
			If CLS_x1_number <> "" then call write_variable_in_case_note("* XFERed to " & CLS_x1_number & ".")
			call write_variable_in_case_note("---")
			call write_variable_in_case_note(worker_signature & ", via automated script.")
		Else
			'Displaying results in Developer Mode
			case_note_box = MsgBox("This case would get case noted if developer mode wasn't on." & worker_signature, vbOKCancel)
			If case_note_box = vbCancel then stopscript
		End if
	' ... or, if the case is not allowed to be XFER'd and it is not privileged, we give it the following case note ...
	ELSEIF privileged_status <> TRUE AND INAC_scrubber_primary_array(x, 8) = FALSE THEN
		call navigate_to_MAXIS_screen("CASE", "NOTE")
		PF9
		tikl_date = dateadd("M", 4, (MAXIS_footer_month & "/01/" & MAXIS_footer_year))
		last_rein_date = dateadd("D", -1, tikl_date)
		IF developer_mode = False THEN
			CALL write_variable_in_case_note("-----ALL PROGRAMS INACTIVE-----")
			CALL write_variable_in_case_note("* Not transfering to CLOSED CASES")
			IF mmis_mode = FALSE THEN CALL write_variable_in_CASE_NOTE("* CL closed on HC for " & inac_month & " but MMIS check bypassed.")
			IF mmis_mode = TRUE THEN CALL write_variable_in_case_note("* HC closed for no-or-incomplete renewal. Last HC REIN Date: " & last_rein_date)
			CALL write_variable_in_case_note("---")
			CALL write_variable_in_case_note(worker_signature)

			'TIKL'ing for 4 months
			IF mmis_mode = TRUE AND closure_reason = "NO REVIEW" THEN
				CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
				CALL create_maxis_friendly_date(tikl_date, 0, 5, 18)
				EMWriteScreen ("IF CASE IS INACTIVE TRANSFER TO CLOSED - " & CLS_x1_number), 9, 3
				transmit
				PF3
			'Adding TIKL to remind worker to check case. This is if MMIS was bypassed.
			ELSEIF mmis_mode = FALSE AND DAILS_out = FALSE THEN
				CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
				CALL create_maxis_friendly_date(date, 0, 5, 18)
				EMWriteScreen ("HC CLOSED BUT MMIS NOT CHECKED. REVIEW TO DETERMINE IF TO SEND TO " & CLS_x1_number), 9, 3
				transmit
				PF3
			END IF

		ELSE
			'Displaying results in developer mode.
			IF INAC_scrubber_primary_array(x, 8) = false THEN MsgBox ("The script would case note the last date to REIN is " & last_rein_date & " and then TIKL to XFER to CLS on " & tikl_date)
			IF mmis_mode = FALSE THEN MsgBox ("Not XFERing case to CLS because MA closed for " & inac_month)
		END IF
	END IF
Next

'If there's no CLS_x1_number, it ends here.
If CLS_x1_number = "" then
	MsgBox 	"Success!"  & vbNewLine & _
			vbNewLine &_
			"The cases that have HC open in MMIS, have unresolved IEVS, or have DAILs generated, did not get case noted. Some of these cases may be discrepancies or may be MCRE or active IMA cases." & vbNewLine & _
			vbNewLine & _
			"A Word document has been created, containing active claims as well as cases with ABPS panels requiring update. If you have questions about these procedures, see a supervisor." & vbNewLine & _
			vbNewLine & _
			"Please note that this script normally XFERs cases to a ''CLS'' account (such as x100CLS), which is used by many agencies to store closed cases. This makes certain functions simpler in an agency. Your agency has not configured such an account for script usage, so it will stop now. If you have additional questions, consult an alpha user."
	script_end_procedure("")
End if

'This do...loop transfers the cases to the CLS_x1_number.
For x = 0 to total_cases
	''Grabs MAXIS_case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
	'MAXIS_case_number = INAC_scrubber_primary_array(x, 0)
	'DAILS_out = INAC_scrubber_primary_array(x, 4)
	'MMIS_status = INAC_scrubber_primary_array(x, 5)
	'privileged_status = INAC_scrubber_primary_array(x, 7)

	'Gets back to SELF (SPEC/XFER gets wonky sometimes, this is safer than using the function)
	back_to_SELF

	'If it isn't privileged, DAILS aren't out there, and MMIS contains no info on this case (or an IMA case), then it'll SPEC/XFER
	If INAC_scrubber_primary_array(x, 8) = TRUE THEN
		EMWriteScreen "SPEC", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen INAC_scrubber_primary_array(x, 0), 18, 43
		EMWriteScreen "XFER", 21, 70
		transmit
		If developer_mode = False then
			EMWriteScreen "x", 7, 16
			transmit
			PF9
			EMWriteScreen CLS_x1_number, 18, 61
			transmit

		Else
			Msgbox INAC_scrubber_primary_array(x, 0) & " transfered to " & CLS_x1_number    'leading a messagebox to show developer what case is being transferred and to where. This pause insures loop is operating correctly.
			objExcel.Cells(x+2, 10).Value = "TRANSFERRED"
		End if
	End if
Next

'Notifies the worker of the success
MsgBox("Success!"  & vbNewLine & _
		vbNewLine &_
		"The cases that have HC open in MMIS, have unresolved IEVS, or have DAILs generated, are still in your REPT/INAC. Some of these cases may be discrepancies or may be MCRE or active IMA cases. Check each one of these manually in MMIS and CCOL/CLIC or process IEVS using TE0019.164 before sending to " & CLS_x1_number & "." & vbNewLine & _
		vbNewLine & _
		"A Word document has been created, containing active claims as well as cases with ABPS panels requiring update. If you have questions about these procedures, see a supervisor.")

script_end_procedure("")
