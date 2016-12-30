'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TYMA TIKLER.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
IF TYMA_TIKL_ALL_AT_ONCE = TRUE THEN STATS_manualtime = 136        'manual run time in seconds for TIKLING all at once
IF TYMA_TIKL_ALL_AT_ONCE = FALSE THEN STATS_manualtime = 30        'manual run time in seconds for TIKLING as you go (1st TIKL)
STATS_denomination = "C"		'C is for case
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

	'DIALOGS--------------------------------------------------------------------------------------------------------------------
	BeginDialog TYMA_tikler, 0, 0, 261, 200, "TYMA/TMA TIKLer"
	EditBox 65, 5, 65, 15, MAXIS_case_number
	EditBox 190, 65, 20, 15, start_month
	EditBox 225, 65, 25, 15, start_year
	ButtonGroup ButtonPressed
		PushButton 185, 90, 65, 20, "Calculate Dates", calculate_button
	EditBox 100, 25, 80, 15, first_quart_send
	EditBox 100, 45, 80, 15, second_quart_send
	EditBox 100, 65, 80, 15, second_quart_due
	EditBox 100, 85, 80, 15, third_quart_send
	EditBox 100, 105, 80, 15, third_quart_due
	EditBox 100, 125, 80, 15, TYMA_close
	EditBox 80, 170, 60, 15, worker_signature
	ButtonGroup ButtonPressed
		OkButton 150, 175, 50, 15
		CancelButton 205, 175, 50, 15
	Text 190, 35, 55, 20, "TYMA/TMA start month/year"
	Text 215, 70, 5, 15, "/"
	Text 5, 30, 90, 10, "1st Quarter form send date"
	Text 5, 90, 90, 10, "3rd Quarter form send date"
	Text 5, 50, 90, 10, "2nd Quarter form send date"
	Text 5, 125, 80, 10, "TYMA closing reminder"
	Text 5, 70, 90, 10, "2nd Quarter form due date"
	Text 5, 110, 90, 10, "3rd Quarter form due date"
	Text 10, 145, 235, 20, "Once you click OK the script will TIKL for the events above on the dates above. Then you can use the DAIL scrubber to resolve the TIKLs."
	Text 10, 10, 55, 10, "Case Number: "
	Text 10, 175, 60, 10, "Worker Signature: "
	EndDialog

	BeginDialog TYMA_tikler_full, 0, 0, 141, 130, "TYMA/TMA TIKLer"
	EditBox 60, 5, 65, 15, MAXIS_case_number
	EditBox 5, 55, 20, 15, start_month
	EditBox 40, 55, 25, 15, start_year
	EditBox 75, 90, 60, 15, worker_signature
	ButtonGroup ButtonPressed
		OkButton 15, 110, 50, 15
		CancelButton 70, 110, 50, 15
	Text 5, 30, 55, 20, "TYMA/TMA start month/year"
	Text 30, 60, 5, 15, "/"
	Text 5, 10, 55, 10, "Case Number: "
	Text 5, 95, 60, 10, "Worker Signature: "
	Text 75, 25, 60, 60, "Script is being run away from DAIL. This will TIKL for 2nd Quarter report form to be sent based on TYMA start month. "
	EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
EMConnect""

'Script runs one of two ways
' TIKLS all at once: Script will create all of the TIKLS in one burst at the start of TYMA.
' TIKLS as you go: Script will create the first TIKL then the worker will use the DAIL Scrubber to create the additional TIKLS.
'The divide will be based on the following variable in the GLOBAL VARIABLES file TYMA_TIKL_ALL_AT_ONCE, the variable will be TRUE/FALSE and restrict an agency to once or the other.

'FIRST HALF------------------------------------------------------------------------------------------------------------------------------------
IF TYMA_TIKL_ALL_AT_ONCE = TRUE THEN    'This section will be dedicated to TIKLing all at once.
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)

	Do
		err_msg = ""
		Do
			dialog TYMA_tikler
			cancel_confirmation
			If buttonpressed = calculate_button Then     'calculate button will calculate the TYMA dates if it is pressed, uses the MM/YY TYMA start date entered by worker.
				If start_month = "" or start_year = "" THEN    'safeguard in case someone clicks calculate without entering a starter month/year
					Msgbox "Please enter a TYMA start month/year in MM YY format then click the calculate button."
				ELSE
					If len(start_month) < 2 THEN start_month = 0 & start_month    '
					If len(start_year) > 2 THEN start_year = right(start_year, 2)
					TYMA_start_date = start_month & "/01/" & start_year
					first_quart_send = DatePart("m", DateAdd("M", 2, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 3, TYMA_start_date))  'date to send 1st quarter report form
					second_quart_send = DatePart("m", DateAdd("M", 5, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 5, TYMA_start_date))  'date to send 2nd quarter report form
					second_quart_due = DatePart("m", DateAdd("M", 6, TYMA_start_date)) & "/21/" & DatePart("YYYY", DateAdd("M", 6, TYMA_start_date))    'date 2nd quarter report form is due
					third_quart_send = DatePart("m", DateAdd("M", 8, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 8, TYMA_start_date))    'date to send 3rd quarter report form
					third_quart_due = DatePart("m", DateAdd("M", 9, TYMA_start_date)) & "/21/" & DatePart("YYYY", DateAdd("M", 9, TYMA_start_date))		'date 3rd quarter report form is due
					TYMA_close = DatePart("m", DateAdd("M", 11, TYMA_start_date)) & "/01/" & DatePart("YYYY", DateAdd("M", 11, TYMA_start_date))		'date to remind worker of TYMA closure
				End If
			End If
		Loop until buttonpressed = OK
		If MAXIS_case_number = "" THEN err_msg = err_msg & Vbcr & "You must enter a case number."     'error message handling builds an error message based on items left off of the dialog.
		If isdate(first_quart_send) = False THEN err_msg = err_msg & vbcr & "1st quarter send date is invalid"
		If isdate(second_quart_send) = False THEN err_msg = err_msg & vbcr & "2nd quarter send date is invalid"
		If isdate(second_quart_due) = False THEN err_msg = err_msg & vbcr & "2nd quarter due date is invalid"
		If isdate(third_quart_send) = False THEN err_msg = err_msg & vbcr & "3rd quarter send date is invalid"
		If isdate(third_quart_due) = False THEN err_msg = err_msg & vbcr & "3rd quarter due date is invalid"
		If isdate(TYMA_close) = False THEN err_msg = err_msg & vbcr & "TYMA end date warning is invalid"
		If worker_signature = "" THEN err_msg = err_msg & Vbcr & "You must enter a worker signature."
		IF err_msg <> "" THEN msgbox err_msg
	Loop until err_msg = ""

	call check_for_MAXIS(false)

	'WRITING THE TIKLS------------------------------------------------------------------------------------------------------------------------------------------------------
	'TIKLS TO SEND FIRST FORM AND DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(first_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 1st Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	'TIKLS TO SEND SECOND FORM AND DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(second_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 2nd Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(second_quart_due, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~2nd Quarterly Report form is now due. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	'TIKLS TO SEND THIRD FORM AND DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(third_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 3rd Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(third_quart_due, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~3rd Quarterly Report form is now due. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	'TIKLS TO SEND ER FORM AND DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(TYMA_close, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~TYMA ending " & Dateadd("m", 12, TYMA_start_date) &  ", take appropriate action. TYMA started " & TYMA_start_date & ". This TIKL was generated via script.")
	Transmit
	PF3


	'THE CASE NOTE PORTION------------------------------------------------------------------------------------------------------
	start_a_blank_CASE_NOTE
	call write_variable_in_CASE_NOTE("***TYMA TIKLS have been created***")
	Call write_bullet_and_variable_in_case_note("TYMA Start date", TYMA_start_date)
	call write_variable_in_CASE_NOTE("The following TIKLs were created by an automated script. When processing these TIKLs follow current procedures.")
	call write_variable_in_CASE_NOTE("---")
	call write_bullet_and_variable_in_case_note("Send 1st Quarter Form", first_quart_send)
	call write_bullet_and_variable_in_case_note("Send 2nd Quarter Form", second_quart_send)
	call write_bullet_and_variable_in_case_note("2nd Quarter Form due", second_quart_due)
	call write_bullet_and_variable_in_case_note("Send 3rd Quarter Form", third_quart_send)
	call write_bullet_and_variable_in_case_note("3rd Quarter Form due", third_quart_due)
	call write_bullet_and_variable_in_case_note("TYMA closure notice", TYMA_close)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	script_end_procedure("Success! TIKLS have been created for TYMA period.")
END If

'SECOND HALF------------------------------------------------------------------------------------------------------------------------------------
IF TYMA_TIKL_ALL_AT_ONCE = FALSE or TYMA_TIKL_ALL_AT_ONCE = "" THEN  'This portion will just add the first TIKL then the worker can use the DAIL scrubber on future TIKLS.
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	Do
		err_msg = ""
		dialog TYMA_tikler_full
		cancel_confirmation
		If MAXIS_case_number = "" THEN err_msg = err_msg & Vbcr & "You must enter a case number."
		If worker_signature = "" THEN err_msg = err_msg & Vbcr & "You must enter a worker signature."
		IF err_msg <> "" THEN msgbox err_msg
	Loop until err_msg = ""

	'Formatting and calculating date for first TIKL in the cycle.
	If len(start_month) < 2 THEN start_month = 0 & start_month    '
	If len(start_year) > 2 THEN start_year = right(start_year, 2)
	TYMA_start_date = start_month & "/01/" & start_year
	first_quart_send = DatePart("m", DateAdd("M", 2, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 3, TYMA_start_date))   'date to send 1st quarter report form

	'TIKLS TO SEND FIRST FORM AND DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	Call create_MAXIS_friendly_date(first_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 1st Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has added a TIKL to send the 1st quarter report form on " & first_quart_send)
END IF
