'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - INAC SCRUBBER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 169                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'THE FOLLOWING VARIABLE IS DYNAMICALLY DETERMINED BY THE PRESENCE OF DATA IN CLS_x1_number. IT WILL BE ADDED DYNAMICALLY TO THE DIALOG BELOW.
If CLS_x1_number <> "" then CLS_dialog_string = "**This script will XFER cases in REPT/INAC to " & CLS_x1_number & ".**"

'THE SCRIPT WILL RUN ONE OF TWO WAYS DEPENDING ON WHAT IS ENTERED INTO GLOBAL VARIABLES.
'FOR AGENCIES THAT WANT TO KEEP INACTIVE MAGI CASES THAT CLOSED FOR NO RENEWAL IN THE CURRENT WORKER'S
'NUMBER FOR 4 MONTHS THE GLOBAL VARIABLE MAGI_cases_closed_four_month_TIKL_no_XFER MUST BE SET TO TRUE. 
'IF SET TO FALSE THE SCRIPT WILL TRANSFER ALL ALLOWABLE INACTIVE CASES TO CLS OR WHEREEVER AGENCY SETS IN GLOBAL VARIABLES.

'MAGI CASES INACTIVE FOR RENENWALS WILL BE HELD ON TO FOR 4 MONTHS--------------------------------------------------------------------------------------------------------------------------------------
IF MAGI_cases_closed_four_month_TIKL_no_XFER = TRUE THEN
	
	'DIALOGS----------------------------------------------------------------------------------------------------
	'NOTE: this dialog uses a dynamic CLS_dialog_string variable. As such, it can't be directly edited in dialog editor.
	BeginDialog INAC_scrubber_dialog, 0, 0, 206, 162, "INAC scrubber dialog"
	EditBox 80, 80, 80, 15, worker_signature
	EditBox 145, 100, 60, 15, worker_number
	EditBox 55, 120, 35, 15, footer_month
	EditBox 145, 120, 35, 15, footer_year
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
	
	'THE SCRIPT----------------------------------------------------------------------------------------------------
	EMConnect ""
	
	'Shows the dialog, requires 7 digits for worker number, a worker signature, a footer_month and year. Contains developer mode to bypass case noting and XFERing.
	Do
		Do
			Do
				Dialog INAC_scrubber_dialog
				If buttonpressed = cancel then stopscript
				If worker_signature = "UUDDLRLRBA" then 
					MsgBox "Developer mode enabled. Will bypass XFER and case note functions."
					worker_signature = ""
					developer_mode = True
				Else
					developer_mode = False
				End if
				If len(worker_number) <> 7 then MsgBox("Your worker number is not 7 digits. Please try again. Type the whole worker number.")
			Loop until len(worker_number) = 7
			If worker_signature = "" and developer_mode = False then MsgBox "You must sign your case notes!"	'Must sign case notes, or be in developer mode which does not case note.
		Loop until worker_signature <> "" or developer_mode = True
		If footer_month = "" or footer_year = "" then MsgBox "You must provide a footer month and year!"
	Loop until footer_month <> "" and footer_year <> ""
	
	
	'Converts system/footer month and year to a MAXIS-appropriate number, for validation
	current_system_month = DatePart("m", Now)
	If len(current_system_month) = 1 then current_system_month = "0" & current_system_month
	current_system_year = DatePart("yyyy", Now) - 2000
	If len(footer_month) <> 2 or isnumeric(footer_month) = False or footer_month > 13 or len(footer_year) <> 2 or isnumeric(footer_year) = False then script_end_procedure("Your footer month and year must be 2 digits and numeric. The script will now stop.")
	footer_month_first_day = footer_month & "/01/" & footer_year
	date_compare = datediff("d", footer_month_first_day, date)
	If date_compare < 0 then script_end_procedure("You appear to have entered a future month and year. The script will now stop.")
	If cint(current_system_month) = cint(footer_month) and cint(footer_year) = cint(current_system_year) AND developer_mode = False then script_end_procedure("Do not use this script in the current footer month. These cases need to be in your REPT/INAC for 30 days. The script will now stop.")
	
	'Warning message before executing
	warning_message = MsgBox(	"Worker: " & worker_number & vbCr & _
								"Footer month/year: " & footer_month & "/" & footer_year & vbCr & _
								vbCr & _
								"This script will case note EACH case on the above REPT/INAC, in the selected footer month, and XFER to " & CLS_x1_number & ", under the following conditions:" & vbCr & _
								"   " & chr(183) & " Case has no open HC on this case number. " & vbCr & _
								"   " & chr(183) & " Case has no open IMA. " & vbCr & _
								"   " & chr(183) & " Case has no messages currently on the DAIL. " & vbCr & _
								"   " & chr(183) & " Case is a closure, and not a denial. For denials, use ''Denied progs''. " & vbCr & _
								vbCr & _
								"This script will also generate a Word document with the following info from the entire caseload: " & vbCr & _
								"   " & chr(183) & " CCOL/CLIC information. " & vbCr & _
								"   " & chr(183) & " Good cause ABPS status. " & vbCr & _
								vbCr & _
								"It requires the use of MDHS for your state systems log-on, as it needs to check MMIS. Also, it only runs in the month before the current footer month (or any month prior)." & vbCr & _
								vbCr & _
								"Please press OK to continue, or cancel to exit the script.", vbOKCancel)
	If warning_message = vbCancel then stopscript
	
	'Connects to MAXIS
	EMConnect ""
	
	'It sends an enter to force the screen to refresh, in order to check for a password prompt.
	MAXIS_check_function
	
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
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	EMWriteScreen "INAC", 21, 70
	transmit
	
	'Checks to make sure the selected worker number is the default. If not it will navigate to that person.
	EMReadScreen worker_number_check, 3, 21, 16
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
		ObjExcel.Cells(1, 9).Value = "MAGI?"
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
			EMReadScreen case_number, 8, MAXIS_row, 3          'First it reads the case number, name, date they closed, and the APPL date.
			EMReadScreen client_name, 25, MAXIS_row, 14
			EMReadScreen inac_date, 8, MAXIS_row, 49
			EMReadScreen appl_date, 8, MAXIS_row, 39
			case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each. 
			client_name = Trim(client_name)
			inac_date = Trim(inac_date)
			appl_date = Trim(appl_date)
			If appl_date <> inac_date then                     'Because if the two dates equal each other, then this is a denial and not a case closure.
				
				'Adds case info to an array. Uses tildes to differentiate the case_number, client_name, and inac_date. Uses vert lines to differentiate entries. Will be fleshed out later.
				If INAC_info_array = "" then
					INAC_info_array = case_number & "~" & client_name & "~" & inac_date
				Else
					INAC_info_array = INAC_info_array & "|" & case_number & "~" & client_name & "~" & inac_date
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
	
	'Assigns info to the array. If developer_mode is on, it'll also add to an Excel spreadsheet
	For x = 0 to total_cases
		interim_array = split(INAC_info_array(x), "~")			'This is a temporary array, and is always three objects (case_number, client_name, INAC_date)
		INAC_scrubber_primary_array(x, 0) = interim_array(0)	'The case_number
		INAC_scrubber_primary_array(x, 1) = interim_array(1)	'The client_name
		INAC_scrubber_primary_array(x, 2) = interim_array(2)	'The inac_date
		If developer_mode = True then 
			ObjExcel.Cells(x + 2, 1).Value = INAC_scrubber_primary_array(x, 0)
			ObjExcel.Cells(x + 2, 2).Value = INAC_scrubber_primary_array(x, 1)
			ObjExcel.Cells(x + 2, 3).Value = INAC_scrubber_primary_array(x, 2)
		End if
	Next
	
	'Navigates to CCOL/CLIC
	EMWriteScreen "CCOL", 20, 22
	EMWriteScreen "CLIC", 20, 70
	transmit
	
	
	'Grabs any claims due for each case. Adds to Excel if developer_mode = True
	For x = 0 to total_cases
		case_number = INAC_scrubber_primary_array(x, 0)
		EMWriteScreen "________", 4, 8
		EMWriteScreen case_number, 4, 8
		transmit
		EMReadScreen claims_due, 10, 19, 58
		INAC_scrubber_primary_array(x, 3) = claims_due
		If developer_mode = True then ObjExcel.Cells(x + 2, 4).Value = claims_due
	Next
	
	'Entering claims into the Word doc
	For x = 0 to total_cases
		'Grabbing the case_number, client_name, and claims_due from the array
		case_number = INAC_scrubber_primary_array(x, 0)
		client_name = INAC_scrubber_primary_array(x, 1)
		claims_due = INAC_scrubber_primary_array(x, 3)
		
		'If there's a claim due, it'll add to the Word doc
		If claims_due <> 0 then 
			objselection.typetext case_number & ": " & client_name & "; amount due: $" & claims_due
			objselection.TypeParagraph()
		End if
	Next
	
	'Navigating to the DAIL (Goes back to self as there are issues in CCOL/CLIC with the direct navigate_to_screen)
	back_to_SELF
	call navigate_to_screen("DAIL", "DAIL")
	
	
	
	'This checks the DAIL for messages, sends a variable to the array. We don't transfer cases with DAIL messages. (True for "has DAIL", False for "doesn't have DAIL")
	For x = 0 to total_cases
		case_number = INAC_scrubber_primary_array(x, 0)		'Grabbing case number
		EMWriteScreen "________", 20, 38
		EMWriteScreen case_number, 20, 38
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
		case_number = INAC_scrubber_primary_array(x, 0)	
		client_name = INAC_scrubber_primary_array(x, 1)	
		
		'Gets to MEMB
		call navigate_to_screen("STAT", "MEMB")
		
		'Checks to make sure we're past SELF. If we aren't, it'll save that the case is privileged (most likely cause) in the array.
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then 
			INAC_scrubber_primary_array(x, 7) = True		'If it's privileged, it won't get past SELF. If so, it'll enter a True value in the array.
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
			call navigate_to_screen("STAT", "ABPS")
			EMReadScreen good_cause_check, 1, 5, 47
			If good_cause_check = "P" then
				objselection.typetext case_number & ", " & client_name
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
	attn
	EMReadScreen MMIS_A_check, 7, MMIS_MDHS_row, 15
	IF MMIS_A_check = "RUNNING" then 
		EMSendKey MMIS_number
		transmit
	End if
	IF MMIS_A_check <> "RUNNING" then 
		attn
		EMConnect "B"
		attn
		EMReadScreen MMIS_B_check, 7, MMIS_MDHS_row, 15
		If MMIS_B_check <> "RUNNING" then 
			script_end_procedure("MMIS does not appear to be running. This script will now stop.")
		End if
		If MMIS_B_check = "RUNNING" then 
			EMSendKey MMIS_number
			transmit
		End if
	End if
	
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
	
		INAC_scrubber_primary_array(x, 8) = False
	
			
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
					IF elig_type = "AX" OR elig_type = "AA" OR elig_type = "CB" OR elig_type = "CK" OR elig_type = "CX" OR elig_type = "PX" THEN INAC_scrubber_primary_array(x, 8) = True
					If elig_end_date = "99/99/99" then 
						INAC_scrubber_primary_array(x, 5) = True
					Else		'Allows for cases that are closing next month
						If datediff("m", elig_end_date, now) < 1 and (isnumeric(MMIS_case_number) = False or MMIS_case_number = MAXIS_case_number) then 
							INAC_scrubber_primary_array(x, 5) = True
						End if
					End if
				End if
				PF6
			End if
		Next
		If INAC_scrubber_primary_array(x, 5) <> True then INAC_scrubber_primary_array(x, 5) = False		'Sets this after the others, so that it doesn't refresh each loop.
		If developer_mode = True then 
			ObjExcel.Cells(x + 2, 6).Value = INAC_scrubber_primary_array(x, 5)		'Writes MMIS status to array when developer_mode is on
			ObjExcel.Cells(x + 2, 9).Value = INAC_scrubber_primary_array(x, 8)
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
	
	'Header for the MMIS discrepancies section of the doc
	objselection.typetext "Case numbers with MMIS discrepancies: "
	objselection.TypeParagraph()
	objselection.TypeParagraph()
	
	'This do...loop updates case notes for all of the cases that don't have DAIL messages or cases still open in MMIS
	For x = 0 to total_cases
		'Grabs case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
		case_number = INAC_scrubber_primary_array(x, 0)
		DAILS_out = INAC_scrubber_primary_array(x, 4)
		MMIS_status = INAC_scrubber_primary_array(x, 5)
		privileged_status = INAC_scrubber_primary_array(x, 7)
	
		'Adds the case number to word doc if MMIS is active
		If MMIS_status = true Then
			objselection.typetext case_number
			objselection.TypeParagraph()
		End If
	
		'Checking to determine that the client is a MAGI that closed for no or incomplete review. If that is the case, then the script does not transfer the client to CLS
		IF INAC_scrubber_primary_array(x, 8) = True THEN
			CALL navigate_to_screen("CASE", "CURR")
			EMWriteScreen "X", 4, 9
			transmit
	
			EMWriteScreen "MA", 3, 19
			transmit
	
			EMReadScreen closure_reason, 9, 8, 60
			EMReadScreen closure_date, 5, 8, 28
			EMReadScreen inac_month, 5, 20, 54
			inac_month = replace(inac_month, " ", "/")
			IF closure_reason = "NO REVIEW" AND inac_month = closure_date THEN 
				INAC_scrubber_primary_array(x, 8) = True
				objSelection.typetext case_number & ": case has MAGI HC client(s) that closed for incomplete or no review."
				objSelection.TypeParagraph()
			ELSE
				INAC_scrubber_primary_array(x, 8) = False
			END IF
		END IF
	
		IF developer_mode = True THEN ObjExcel.Cells(x + 2, 9).Value = INAC_scrubber_primary_array(x, 8)
		
		back_to_self
		'If it isn't privileged, DAILS aren't out there, and MMIS contains no info on this case (or an IMA case), then it'll case note.
		If privileged_status <> True and DAILS_out = False and MMIS_status = False  AND INAC_scrubber_primary_array(x, 8) = False then
			call navigate_to_screen("CASE", "NOTE")
			PF9
			If developer_mode = False then
				call write_new_line_in_case_note("--------------------Case is closed--------------------")
				call write_new_line_in_case_note("* Reviewed closed case for claims via automated script.")
				If CLS_x1_number <> "" then call write_new_line_in_case_note("* XFERed to " & CLS_x1_number & ".")
				call write_new_line_in_case_note("---")
				call write_new_line_in_case_note(worker_signature & ", via automated script.")
			Else
				'case_note_box = MsgBox("This case would get case noted if developer mode wasn't on." & worker_signature, vbOKCancel)
				If case_note_box = vbCancel then stopscript
			End if
		End if
		If privileged_status <> True and DAILS_out = False and MMIS_status = False AND INAC_scrubber_primary_array(x, 8) = True then
			call navigate_to_screen("CASE", "NOTE")
			PF9
			tikl_date = dateadd("M", 4, (footer_month & "/01/" & footer_year))
			last_rein_date = dateadd("D", -1, tikl_date)
			IF developer_mode = False THEN
				CALL write_new_line_in_case_note("-----ALL PROGRAMS INACTIVE-----")
				CALL write_new_line_in_case_note("* Not transfering to Closed Cases because of current MAGI rules")
				CALL write_new_line_in_case_note("* Last HC REIN Date for MAGI client: " & last_rein_date)
				CALL write_new_line_in_case_note("---")
				CALL write_new_line_in_case_note(worker_signature)
	
				CALL navigate_to_screen("DAIL", "WRIT")
				CALL create_maxis_friendly_date(tikl_date, 0, 5, 18)
				EMWriteScreen ("IF CASE IS INACTIVE TRANSFER TO CLOSED - " & CLS_x1_number), 9, 3
				transmit
				PF3
			ELSE
				'MsgBox ("The script would case note the last date to REIN is " & last_rein_date & " and then TIKL to XFER to CLS on " & tikl_date)
			END IF
		END IF
	Next
	
	CLS_x1_number = "X102CLS"
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
		'Grabs case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
		case_number = INAC_scrubber_primary_array(x, 0)
		DAILS_out = INAC_scrubber_primary_array(x, 4)
		MMIS_status = INAC_scrubber_primary_array(x, 5)
		privileged_status = INAC_scrubber_primary_array(x, 7)
	
		'Gets back to SELF (SPEC/XFER gets wonky sometimes, this is safer than using the function)
		back_to_SELF
		
		'If it isn't privileged, DAILS aren't out there, and MMIS contains no info on this case (or an IMA case), then it'll SPEC/XFER
		If privileged_status <> True and DAILS_out = False and MMIS_status = False AND INAC_scrubber_primary_array(x, 8) = False then
			EMWriteScreen "SPEC", 16, 43
			EMWriteScreen "________", 18, 43
			EMWriteScreen case_number, 18, 43
			EMWriteScreen "XFER", 21, 70
			transmit
			If developer_mode = False then
				EMWriteScreen "x", 7, 16
				transmit
				PF9
				EMWriteScreen CLS_x1_number, 18, 61		
				transmit
			Else
				'MsgBox "Case would be XFERed to " & CLS_x1_number & ""
			End if
		End if
	Next
	
	'Notifies the worker of the success
	MsgBox 	"Success!"  & vbNewLine & _
			vbNewLine &_
			"The cases that have HC open in MMIS, have unresolved IEVS, or have DAILs generated, are still in your REPT/INAC. Some of these cases may be discrepancies or may be MCRE or active IMA cases. Check each one of these manually in MMIS and CCOL/CLIC or process IEVS using TE0019.164 before sending to " & CLS_x1_number & "." & vbNewLine & _
			vbNewLine & _
			"A Word document has been created, containing active claims as well as cases with ABPS panels requiring update. If you have questions about these procedures, see a supervisor."
END IF

'MAGI CASES INACTIVE FOR RENEWAL WILL BE TRANSFERRED TO CLOSED CASES
IF MAGI_cases_closed_four_month_TIKL_no_XFER = FALSE THEN

	'DIALOGS----------------------------------------------------------------------------------------------------
	'NOTE: this dialog uses a dynamic CLS_dialog_string variable. As such, it can't be directly edited in dialog editor.
	BeginDialog INAC_scrubber_dialog, 0, 0, 211, 165, "INAC scrubber dialog"
	EditBox 80, 80, 80, 15, worker_signature
	EditBox 145, 100, 60, 15, worker_number
	EditBox 55, 120, 35, 15, footer_month
	EditBox 145, 120, 35, 15, footer_year
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
	
	'THE SCRIPT----------------------------------------------------------------------------------------------------
	'Shows the dialog, requires 7 digits for worker number, a worker signature, a footer_month and year. Contains developer mode to bypass case noting and XFERing.
	Do
		Do
			Do
				Dialog INAC_scrubber_dialog
				If buttonpressed = cancel then stopscript
				If worker_signature = "UUDDLRLRBA" then
					MsgBox "Developer mode enabled. Will bypass XFER and case note functions."
					worker_signature = ""
					developer_mode = True
				Else
					developer_mode = False
				End if
				If len(worker_number) <> 7 then MsgBox("Your worker number is not 7 digits. Please try again. Type the whole worker number.")
			Loop until len(worker_number) = 7
			If worker_signature = "" and developer_mode = False then MsgBox "You must sign your case notes!"	'Must sign case notes, or be in developer mode which does not case note.
		Loop until worker_signature <> "" or developer_mode = True
		If footer_month = "" or footer_year = "" then MsgBox "You must provide a footer month and year!"
	Loop until footer_month <> "" and footer_year <> ""
	
	'Converts worker number to the MAXIS friendly all-caps
	worker_number = UCase(worker_number)
	
	'Converts system/footer month and year to a MAXIS-appropriate number, for validation
	current_system_month = DatePart("m", Now)
	If len(current_system_month) = 1 then current_system_month = "0" & current_system_month
	current_system_year = DatePart("yyyy", Now) - 2000
	If len(footer_month) <> 2 or isnumeric(footer_month) = False or footer_month > 13 or len(footer_year) <> 2 or isnumeric(footer_year) = False then script_end_procedure("Your footer month and year must be 2 digits and numeric. The script will now stop.")
	footer_month_first_day = footer_month & "/01/" & footer_year
	date_compare = datediff("d", footer_month_first_day, date)
	If date_compare < 0 then script_end_procedure("You appear to have entered a future month and year. The script will now stop.")
	If cint(current_system_month) = cint(footer_month) and cint(footer_year) = cint(current_system_year) then script_end_procedure("Do not use this script in the current footer month. These cases need to be in your REPT/INAC for 30 days. The script will now stop.")
	
	'Warning message before executing
	warning_message = MsgBox(	"Worker: " & worker_number & vbCr & _
								"Footer month/year: " & footer_month & "/" & footer_year & vbCr & _
								vbCr & _
								"This script will case note EACH case on the above REPT/INAC, in the selected footer month, and XFER to " & CLS_x1_number & ", under the following conditions:" & vbCr & _
								"   " & chr(183) & " Case has no open HC on this case number. " & vbCr & _
								"   " & chr(183) & " Case has no open IMA. " & vbCr & _
								"   " & chr(183) & " Case has no messages currently on the DAIL. " & vbCr & _
								"   " & chr(183) & " Case is a closure, and not a denial. For denials, use ''Denied progs''. " & vbCr & _
								vbCr & _
								"This script will also generate a Word document with the following info from the entire caseload: " & vbCr & _
								"   " & chr(183) & " CCOL/CLIC information. " & vbCr & _
								"   " & chr(183) & " Good cause ABPS status. " & vbCr & _
								vbCr & _
								"It requires the use of MDHS for your state systems log-on, as it needs to check MMIS. Also, it only runs in the month before the current footer month (or any month prior)." & vbCr & _
								vbCr & _
								"Please press OK to continue, or cancel to exit the script.", vbOKCancel)
	If warning_message = vbCancel then stopscript
	
	'Connects to MAXIS
	EMConnect ""
	
	'It sends an enter to force the screen to refresh, in order to check for a password prompt.
	Call check_for_MAXIS(True)
	
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
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
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
			EMReadScreen case_number, 8, MAXIS_row, 3          'First it reads the case number, name, date they closed, and the APPL date.
			EMReadScreen client_name, 25, MAXIS_row, 14
			EMReadScreen inac_date, 8, MAXIS_row, 49
			EMReadScreen appl_date, 8, MAXIS_row, 39
			case_number = Trim(case_number)                    'Then it trims the spaces from the edges of each.
			client_name = Trim(client_name)
			inac_date = Trim(inac_date)
			appl_date = Trim(appl_date)
			If appl_date <> inac_date then                     'Because if the two dates equal each other, then this is a denial and not a case closure.
	
				'Adds case info to an array. Uses tildes to differentiate the case_number, client_name, and inac_date. Uses vert lines to differentiate entries. Will be fleshed out later.
				If INAC_info_array = "" then
					INAC_info_array = case_number & "~" & client_name & "~" & inac_date
				Else
					INAC_info_array = INAC_info_array & "|" & case_number & "~" & client_name & "~" & inac_date
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
	
	'Declares INAC_scrubber_primary_array_2, redims it to be the size needed for our total amount of cases. THIS NEEDS TO BE DIFFERENT THAN ABOVE BECAUSE YOU CAN'T DIM SAME VARIABLE TWICE.
	Dim INAC_scrubber_primary_array_2()
	ReDim INAC_scrubber_primary_array_2(total_cases, 7)
	
	'Assigns info to the array. If developer_mode is on, it'll also add to an Excel spreadsheet
	For x = 0 to total_cases
		interim_array = split(INAC_info_array(x), "~")			'This is a temporary array, and is always three objects (case_number, client_name, INAC_date)
		INAC_scrubber_primary_array_2(x, 0) = interim_array(0)	'The case_number
		INAC_scrubber_primary_array_2(x, 1) = interim_array(1)	'The client_name
		INAC_scrubber_primary_array_2(x, 2) = interim_array(2)	'The inac_date
		If developer_mode = True then
			ObjExcel.Cells(x + 2, 1).Value = INAC_scrubber_primary_array_2(x, 0)
			ObjExcel.Cells(x + 2, 2).Value = INAC_scrubber_primary_array_2(x, 1)
			ObjExcel.Cells(x + 2, 3).Value = INAC_scrubber_primary_array_2(x, 2)
		End if
	Next
	
	'Navigates to CCOL/CLIC
	EMWriteScreen "CCOL", 20, 22
	EMWriteScreen "CLIC", 20, 70
	transmit
	
	'Grabs any claims due for each case. Adds to Excel if developer_mode = True
	For x = 0 to total_cases
		case_number = INAC_scrubber_primary_array_2(x, 0)
		EMWriteScreen "________", 4, 8
		EMWriteScreen case_number, 4, 8
		transmit
		EMReadScreen claims_due, 10, 19, 58
		INAC_scrubber_primary_array_2(x, 3) = claims_due
		If developer_mode = True then ObjExcel.Cells(x + 2, 4).Value = claims_due
	Next
	
	'Entering claims into the Word doc
	For x = 0 to total_cases
		'Grabbing the case_number, client_name, and claims_due from the array
		case_number = INAC_scrubber_primary_array_2(x, 0)
		client_name = INAC_scrubber_primary_array_2(x, 1)
		claims_due = INAC_scrubber_primary_array_2(x, 3)
	
		'If there's a claim due, it'll add to the Word doc
		If claims_due <> 0 then
			objselection.typetext case_number & ": " & client_name & "; amount due: $" & claims_due
			objselection.TypeParagraph()
		End if
	Next
	
	'Navigating to the DAIL (Goes back to self as there are issues in CCOL/CLIC with the direct navigate_to_MAXIS_screen)
	back_to_SELF
	call navigate_to_MAXIS_screen("DAIL", "DAIL")
	
	'This checks the DAIL for messages, sends a variable to the array. We don't transfer cases with DAIL messages. (True for "has DAIL", False for "doesn't have DAIL")
	For x = 0 to total_cases
		case_number = INAC_scrubber_primary_array_2(x, 0)		'Grabbing case number
		EMWriteScreen "________", 20, 38
		EMWriteScreen case_number, 20, 38
		transmit
		EMReadScreen DAIL_check, 1, 5, 5
		If DAIL_check <> " " then
			INAC_scrubber_primary_array_2(x, 4) = True
		Else
			INAC_scrubber_primary_array_2(x, 4) = False
		End if
		If developer_mode = True then ObjExcel.Cells(x + 2, 5).Value = INAC_scrubber_primary_array_2(x, 4)
		excel_row = excel_row + 1
		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
	Next
	
	'Making the header for the next section of the Word document.
	objselection.TypeParagraph()
	objselection.TypeParagraph()
	objselection.typetext "Cases that need to be REINed, STAT/ABPS updated with an ''N'' code for Good Cause Status, and then reapproved for closure:"
	objselection.TypeParagraph()
	
	'This do...loop goes into STAT, grabs PMIs for MEMB types 01, 02, 03, 04, and 18, and then navigates to ABPS to get that info.
	For x = 0 to total_cases
	
		'Grabbing case number and name for this loop
		case_number = INAC_scrubber_primary_array_2(x, 0)
		client_name = INAC_scrubber_primary_array_2(x, 1)
	
		'Gets to MEMB
		call navigate_to_MAXIS_screen("STAT", "MEMB")
	
		'Checks to make sure we're past SELF. If we aren't, it'll save that the case is privileged (most likely cause) in the array.
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			INAC_scrubber_primary_array_2(x, 7) = True		'If it's privileged, it won't get past SELF. If so, it'll enter a True value in the array.
		Else
			INAC_scrubber_primary_array_2(x, 7) = False		'If it gets through, it isn't privileged.
	
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
				INAC_scrubber_primary_array_2(x, 6) = PMI_array		'Writes the PMI array to the main array
			Loop until no_more_MEMBs_check = "ENTER A VALID COMMAND OR PF-KEY"
	
			'Goes to ABPS to check good cause. Good cause will not hang the case from being sent to CLS, as such, it does not get entered in the array (just the Word doc).
			call navigate_to_MAXIS_screen("STAT", "ABPS")
			EMReadScreen good_cause_check, 1, 5, 47
			If good_cause_check = "P" then
				objselection.typetext case_number & ", " & client_name
				objselection.TypeParagraph()
			End if
		End if
		If developer_mode = True then
			ObjExcel.Cells(x + 2, 8).Value = INAC_scrubber_primary_array_2(x, 7)		'Writes privileged status to Excel when developer_mode is on
			ObjExcel.Cells(x + 2, 7).Value = INAC_scrubber_primary_array_2(x, 6)		'Writes PMI array to Excel when developer_mode is on
		End if
		PMI_array = ""		'Clears the variable for the following loop
	Next
	
	'MMIS--------------------------------------------------------------------------------------------------------------
	'The following checks for which screen MMIS is running on.
	attn
	EMReadScreen MMIS_A_check, 7, MMIS_MDHS_row, 15
	IF MMIS_A_check = "RUNNING" then
		EMSendKey MMIS_number
		transmit
	End if
	IF MMIS_A_check <> "RUNNING" then
		attn
		EMConnect "B"
		attn
		EMReadScreen MMIS_B_check, 7, MMIS_MDHS_row, 15
		If MMIS_B_check <> "RUNNING" then
			script_end_procedure("MMIS does not appear to be running. This script will now stop.")
		End if
		If MMIS_B_check = "RUNNING" then
			EMSendKey MMIS_number
			transmit
		End if
	End if
	
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
		MAXIS_case_number = INAC_scrubber_primary_array_2(x, 0)
		PMI_array = INAC_scrubber_primary_array_2(x, 6)
		privileged_status = INAC_scrubber_primary_array_2(x, 7)
	
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
						INAC_scrubber_primary_array_2(x, 5) = True
					End if
				ElseIf MMIS_case_status = "C" or MMIS_case_status = "D" then
					EMReadScreen elig_end_date, 8, 7, 36
					If elig_end_date = "99/99/99" then
						INAC_scrubber_primary_array_2(x, 5) = True
					Else		'Allows for cases that are closing next month
						If datediff("m", elig_end_date, now) < 1 and (isnumeric(MMIS_case_number) = False or MMIS_case_number = MAXIS_case_number) then
							INAC_scrubber_primary_array_2(x, 5) = True
						End if
					End if
				End if
				PF6
			End if
		Next
		If INAC_scrubber_primary_array_2(x, 5) <> True then INAC_scrubber_primary_array_2(x, 5) = False		'Sets this after the others, so that it doesn't refresh each loop.
		If developer_mode = True then ObjExcel.Cells(x + 2, 6).Value = INAC_scrubber_primary_array_2(x, 5)		'Writes MMIS status to array when developer_mode is on
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
			script_end_procedure("MAXIS does not appear to be running. This script will now stop.")
		Else
			EMSendkey MAXIS_number
			transmit
		End if
	End if
	
	'Header for the MMIS discrepancies section of the doc
	objselection.typetext "Case numbers with MMIS discrepancies: "
	objselection.TypeParagraph()
	objselection.TypeParagraph()
	
	'This do...loop updates case notes for all of the cases that don't have DAIL messages or cases still open in MMIS
	For x = 0 to total_cases
	
		'Grabs case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
		case_number = INAC_scrubber_primary_array_2(x, 0)
		DAILS_out = INAC_scrubber_primary_array_2(x, 4)
		MMIS_status = INAC_scrubber_primary_array_2(x, 5)
		privileged_status = INAC_scrubber_primary_array_2(x, 7)
			'Adds the case number to word doc if MMIS is active
		If MMIS_status = true Then
			objselection.typetext case_number
			objselection.TypeParagraph()
		End If
	
		back_to_self
		'If it isn't privileged, DAILS aren't out there, and MMIS contains no info on this case (or an IMA case), then it'll case note.
		If privileged_status <> True and DAILS_out = False and MMIS_status = False then
			call navigate_to_MAXIS_screen("CASE", "NOTE")
			PF9
			If developer_mode = False then
				call write_variable_in_CASE_NOTE("--------------------Case is closed--------------------")
				call write_variable_in_CASE_NOTE("* Reviewed closed case for claims via automated script.")
				If CLS_x1_number <> "" then call write_variable_in_CASE_NOTE("* XFERed to " & CLS_x1_number & ".")
				call write_variable_in_CASE_NOTE("---")
				call write_variable_in_CASE_NOTE(worker_signature & ", via automated script.")
			Else
				case_note_box = MsgBox("This case would get case noted if developer mode wasn't on." & worker_signature, vbOKCancel)
				If case_note_box = vbCancel then stopscript
			End if
		End if
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
		'Grabs case_number, DAIL info (if any messages are unresolved), MMIS_status, and privileged_status from the main array
		case_number = INAC_scrubber_primary_array_2(x, 0)
		DAILS_out = INAC_scrubber_primary_array_2(x, 4)
		MMIS_status = INAC_scrubber_primary_array_2(x, 5)
		privileged_status = INAC_scrubber_primary_array_2(x, 7)
	
		'Gets back to SELF (SPEC/XFER gets wonky sometimes, this is safer than using the function)
		back_to_SELF
	
		'If it isn't privileged, DAILS aren't out there, and MMIS contains no info on this case (or an IMA case), then it'll SPEC/XFER
		If privileged_status <> True and DAILS_out = False and MMIS_status = False then
			EMWriteScreen "SPEC", 16, 43
			EMWriteScreen "________", 18, 43
			EMWriteScreen case_number, 18, 43
			EMWriteScreen "XFER", 21, 70
			transmit
			If developer_mode = False then
				EMWriteScreen "x", 7, 16
				transmit
				PF9
				EMWriteScreen CLS_x1_number, 18, 61
				transmit
			Else
				MsgBox "Case would be XFERed to " & CLS_x1_number & ""
			End if
		End if
	Next
	
	'Notifies the worker of the success
	MsgBox 	"Success!"  & vbNewLine & _
			vbNewLine &_
			"The cases that have HC open in MMIS, have unresolved IEVS, or have DAILs generated, are still in your REPT/INAC. Some of these cases may be discrepancies or may be MCRE or active IMA cases. Check each one of these manually in MMIS and CCOL/CLIC or process IEVS using TE0019.164 before sending to " & CLS_x1_number & "." & vbNewLine & _
			vbNewLine & _
			"A Word document has been created, containing active claims as well as cases with ABPS panels requiring update. If you have questions about these procedures, see a supervisor."
	
	STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

END IF

script_end_procedure("")
