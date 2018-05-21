'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - RETURNED MAIL WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 70                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("04/04/2017", "Corrected bug with missing worker signature.  Added handling for multiple recipient changes to SPEC/WCOM", "David Courtright, St Louis County")
call changelog_update("01/17/2017", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialog--------------------------------------------
BeginDialog returned_mail_dlg, 0, 0, 156, 120, "Returned Mail WCOM"
  EditBox 65, 5, 75, 15, MAXIS_case_number
  EditBox 60, 25, 20, 15, MAXIS_footer_month
  EditBox 130, 25, 20, 15, MAXIS_footer_year
  EditBox 80, 50, 70, 15, contact_request_date
  EditBox 80, 70, 70, 15, contact_due_date
  ButtonGroup ButtonPressed
    OkButton 25, 95, 50, 15
    CancelButton 80, 95, 50, 15
  Text 85, 25, 40, 20, "Footer Year (YY):"
  Text 5, 55, 70, 10, "Contact request date:"
  Text 10, 10, 55, 10, "Case Number: "
  Text 10, 25, 45, 20, "Footer Month (MM):"
  Text 5, 75, 60, 10, "Contact Due date:"
EndDialog

'The script-------------------------------------
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'warning box
Msgbox "Warning: If you have multiple waiting SNAP results this script may be unable to find the most recent one. Please process manually in those instances."

'the dialog
Do
	Do
		dialog returned_mail_dlg
		cancel_confirmation
		If MAXIS_footer_month = "" or MAXIS_footer_year = "" THEN Msgbox "Please fill in footer month and year (MM YY format)."
		If MAXIS_case_number = "" THEN MsgBox "Please enter a case number."
		If worker_signature = "" THEN MsgBox "Please sign your note."
	Loop until MAXIS_footer_month <> "" & MAXIS_footer_year <> ""
Loop until MAXIS_case_number <> ""

'Converting dates into useable forms
If len(MAXIS_footer_month) < 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
If len(MAXIS_footer_year) > 2 THEN MAXIS_footer_year = right(MAXIS_footer_year, 2)

'Navigating to the spec wcom screen
CALL Check_for_MAXIS(false)

back_to_self

Emwritescreen MAXIS_case_number, 18, 43
Emwritescreen MAXIS_footer_month, 20, 43
Emwritescreen MAXIS_footer_year, 20, 46
 'This section will check for whether forms go to AREP and SWKR

 call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
 EMReadscreen forms_to_arep, 1, 10, 45
 call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
 EMReadscreen forms_to_swkr, 1, 15, 63

CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting SNAP notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "FS" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			Transmit
			PF9
			'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
			IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
			row = 4                             'reset row and col for the next search
			col = 1
			EMSearch "SOCWKR", row, col
			IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
			EMWriteScreen "x", 5, 12                                        'We always send notice to client
			IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			transmit                                                        'Transmits to start the memo writing process'
			Emreadscreen fs_wcom_exists, 3, 3, 15
			If fs_wcom_exists <> "   " then script_end_procedure ("It appears you already have a WCOM added to this notice. The script will now end.")
			If program_type = "FS" AND print_status = "Waiting" then
				fs_wcom_writen = true
				'This will write if the notice is for SNAP only
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("Your mail has been returned to our agency. On " & contact_request_date & " you were sent a request for you to contact this agency because of this returned mail. You did not contact the agency by " & contact_due_date & " so your SNAP case has been closed.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				PF4
				PF3
			End if
		End If
	End If
	If fs_wcom_writen = true then Exit Do
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_fs_waiting = true then script_end_procedure("No waiting FS notice was found for the requested month")
script_end_procedure("WCOM has been added to the first found waiting SNAP notice for the month and case selected. Please review the notice.")