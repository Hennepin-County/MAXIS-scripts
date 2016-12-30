'Required for statistical purposes===============================================================================
name_of_script = "DAIL - TYMA SCRUBBER.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 50        'manual run time in seconds
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

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in DAIL, Please navigate to DAIL and run the script again.")  'if the worker is not on a dail message


'The following reads the message in full for the end part (which tells the script which message was selected)
EMReadScreen full_message, 23, 6, 20
IF full_message = "~*~CONSIDER SENDING 1ST" THEN     'script finds 1st TIKL message and moves to take next action
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	EMWritescreen "X", 6, 3
	Transmit
	EMReadScreen MAXIS_case_number, 8, 6, 57
	EMReadScreen TYMA_start_date, 8, 10, 5          'reading TYMA start date to carry it forward
	TYMA_start_date = cdate(TYMA_start_date)
	Back_to_self
	start_a_blank_CASE_NOTE
	call write_variable_in_CASE_NOTE("***TYMA 1st Quarterly Report Form Sent***")
	call write_variable_in_CASE_NOTE("TIKL created to send 2nd Quarterly Report Form")
	call write_variable_in_CASE_NOTE("-Case note and TIKL created by Automated script")
	'TIKLS TO SEND SECOND FORM
	Call navigate_to_MAXIS_screen("dail", "writ")
	second_quart_send = DatePart("m", DateAdd("M", 5, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 5, TYMA_start_date))  'date to send 2nd quarter report form
	Call create_MAXIS_friendly_date(second_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 2nd Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has case noted that 1st Quarter form was sent and added a TIKL to send the 2nd quarter report form on " & second_quart_send)
END IF
IF full_message = "~*~CONSIDER SENDING 2ND" THEN     'script finds 2nd TIKL message and moves to take next action
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	EMWritescreen "X", 6, 3
	Transmit
	EMReadScreen MAXIS_case_number, 8, 6, 57
	EMReadScreen TYMA_start_date, 8, 10, 5          'reading TYMA start date to carry it forward Needs to read 10 digits since after first TIKL the variable gets Cdated/written into a YYYY format
	TYMA_start_date = cdate(TYMA_start_date)
	Back_to_self
	start_a_blank_CASE_NOTE
	call write_variable_in_CASE_NOTE("***TYMA 2nd Quarterly Report Form Sent***")
	call write_variable_in_CASE_NOTE("TIKL created for 2nd Quarterly Report Form due date")
	call write_variable_in_CASE_NOTE("-Case note and TIKL created by Automated script")
	'TIKLS FOR SECOND FORM DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	second_quart_due = DatePart("m", DateAdd("M", 6, TYMA_start_date)) & "/21/" & DatePart("YYYY", DateAdd("M", 6, TYMA_start_date))   'date 2nd quarter report form is due
	Call create_MAXIS_friendly_date(second_quart_due, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~2nd Quarterly Report form is now due. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has case noted that 2nd Quarter form was sent and added a TIKL for return of 2nd quarter report form on " & second_quart_due)
END IF
IF full_message = "~*~2ND QUARTERLY REPORT" THEN     'script finds 3rd TIKL message and moves to take next action
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	EMWritescreen "X", 6, 3
	Transmit
	EMReadScreen MAXIS_case_number, 8, 6, 57
	EMReadScreen TYMA_start_date, 8, 10, 5          'reading TYMA start date to carry it forward Needs to read 10 digits since after first TIKL the variable gets Cdated/written into a YYYY format
	TYMA_start_date = cdate(TYMA_start_date)
	'TIKLS FOR THIRD FORM SEND DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	third_quart_send = DatePart("m", DateAdd("M", 8, TYMA_start_date)) & "/20/" & DatePart("YYYY", DateAdd("M", 8, TYMA_start_date))   'date to send 3rd quarter report form
	Call create_MAXIS_friendly_date(third_quart_send, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~Consider sending 3rd Quarterly Report form. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has added a TIKL to send the 3rd quarter report form on " & third_quart_send)
END IF
IF full_message = "~*~CONSIDER SENDING 3RD" THEN     'script finds 4th TIKL message and moves to take next action
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	EMWritescreen "X", 6, 3
	Transmit
	EMReadScreen MAXIS_case_number, 8, 6, 57
	EMReadScreen TYMA_start_date, 8, 10, 5          'reading TYMA start date to carry it forward Needs to read 10 digits since after first TIKL the variable gets Cdated/written into a YYYY format
	TYMA_start_date = cdate(TYMA_start_date)
	Back_to_self
	start_a_blank_CASE_NOTE
	call write_variable_in_CASE_NOTE("***TYMA 3rd Quarterly Report Form Sent***")
	call write_variable_in_CASE_NOTE("TIKL created for 3rd Quarterly Report Form due date")
	call write_variable_in_CASE_NOTE("-Case note and TIKL created by Automated script")
	'TIKLS FOR SECOND FORM DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	third_quart_due = DatePart("m", DateAdd("M", 9, TYMA_start_date)) & "/21/" & DatePart("YYYY", DateAdd("M", 9, TYMA_start_date))    'date 3rd quarter report form is due
	Call create_MAXIS_friendly_date(third_quart_due, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~3rd Quarterly Report form is now due. TYMA start: " & TYMA_start_date & ",  Take appropriate action. This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has case noted that 3rd Quarter form was sent and added a TIKL for return of 3rd quarter report form on " & third_quart_due)
END IF
IF full_message = "~*~3RD QUARTERLY REPORT" THEN     'script finds 5th TIKL message and moves to take next action
	call MAXIS_case_number_finder(MAXIS_case_number)
	call check_for_MAXIS(false)
	EMWritescreen "X", 6, 3
	Transmit
	EMReadScreen MAXIS_case_number, 8, 6, 57
	EMReadScreen TYMA_start_date, 8, 10, 5          'reading TYMA start date to carry it forward Needs to read 10 digits since after first TIKL the variable gets Cdated/written into a YYYY format
	TYMA_start_date = cdate(TYMA_start_date)
	'TIKLS FOR SECOND FORM DUE DATE
	Call navigate_to_MAXIS_screen("dail", "writ")
	TYMA_close = DatePart("m", DateAdd("M", 11, TYMA_start_date)) & "/01/" & DatePart("YYYY", DateAdd("M", 11, TYMA_start_date))    'date to remind worker to TYMA is closing
	Call create_MAXIS_friendly_date(TYMA_close, 0, 5, 18)
	Transmit
	Call write_variable_in_TIKL("~*~TYMA ending " & Dateadd("m", 12, TYMA_start_date) &  ", take appropriate action. TYMA started " & TYMA_start_date & ". This TIKL was generated via script.")
	Transmit
	PF3
	script_end_procedure("Success! Script has case noted to remind of the end of TYMA.")
END IF

'If the message doesn't match any of the ones above you get this message.
script_end_procedure("A Valid TYMA DAIL message was not found. Please place your cursor over one and try again, or navigate away from DAIL and restart script if trying to create first TIKL.")
