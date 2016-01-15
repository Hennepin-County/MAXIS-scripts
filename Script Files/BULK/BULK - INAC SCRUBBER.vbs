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
