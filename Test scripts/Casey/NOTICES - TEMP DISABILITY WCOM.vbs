'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - TEMP DISABILITY WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("08/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialogs
BeginDialog case_number_dlg, 0, 0, 151, 130, "Temporary Disabled ABAWD's WCOM"
  EditBox 65, 50, 60, 15, MAXIS_case_number
  EditBox 85, 70, 20, 15, Edit5
  EditBox 85, 90, 20, 15, approval_month
  EditBox 105, 90, 20, 15, approval_year
  ButtonGroup ButtonPressed
    OkButton 20, 110, 50, 15
    CancelButton 75, 110, 50, 15
  Text 15, 55, 50, 10, "Case Number:"
  GroupBox 5, 5, 140, 40, "Purpose of script:"
  Text 15, 15, 125, 25, "To explain the temp disability ABAWD time frame to the client in plain language when SNAP is approved. "
  Text 15, 95, 70, 10, "Approval month/year:"
  Text 35, 75, 50, 10, "HH member #:"
EndDialog

'--- -----------------------------------------------------------------------------------------------------------------The script
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
approval_month = CM_plus_one_mo
approval_year = CM_plus_one_yr

Do
    DO
    	err_msg = ""
    	dialog case_number_dlg
    	If ButtonPressed = 0 then StopScript
    	IF MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a case number"
		If IsNumeric(HH_memb) = False or len(HH_memb) <> 2 then err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit HH member number."
    	IF len(approval_month) <> 2 then err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit footer month."
    	IF len(approval_year) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid 2-digit footer year"
    	IF err_msg <> "" THEN msgbox err_msg
    LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("STAT", "DISA")
Call write_value_and_transmit(HH_memb, 20, 76)

'Reading the disa dates
EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
EMReadScreen disa_end_date, 10, 6, 69
disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
disa_end_date = Replace(disa_end_date," ","/")

If disa_end_date = "__/__/____" then script_end_procedure("Household member has an open-ended DISA date. Please update the disability end date to either the day before the renewal is due, or the date on the Medical Opinion Form, whichever is longer.")

disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
If disa_dates = "__/__/____ - __/__/____" then script_end_procedure("Unable to find disability dates for the HH member identified. Please try again.")

Call navigate_to_MAXIS_screen("STAT", "WREG")
Call write_value_and_transmit(HH_memb, 20, 76)
EMReadScreen FSET_code, 2, 8, 50
If FSET_code <> "03" then script_end_procedure("DISA member must be coded with FSET/ABWAD codes of 03/01.")
EMReadScreen ABAWD_code, 2, 13, 50
If ABAWD_code <> "03" then script_end_procedure("DISA member must be coded with FSET/ABWAD codes of 03/01.")

'This section will check for whether forms go to AREP and SWKR
call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
EMReadscreen forms_to_arep, 1, 10, 45
call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
EMReadscreen forms_to_swkr, 1, 15, 63

call navigate_to_MAXIS_screen("spec", "wcom")

EMWriteScreen approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	waiting_check = ""
	EMReadscreen prog_type, 2, read_row, 26
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
		EMSetcursor read_row, 13
		EMSendKey "x"
		Transmit
		pf9
		'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
		IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
		row = 4                             'reset row and col for the next search
		col = 1
		EMSearch "SOCWKR", row, col
		IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
		EMWriteScreen "x", 5, 10                                        'We always send notice to client
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		transmit                                                        'Transmits to start the memo writing process'
		EMSetCursor 03, 15
		CALL write_variable_in_SPEC_MEMO("You are no longer eligible for MFIP because " & MFIP_closure_reason)
		CALL write_variable_in_SPEC_MEMO(" ")
		CALL write_variable_in_SPEC_MEMO("Please contact your worker with any questions. Thank you.")
		PF4
		PF3
		WCOM_count = WCOM_count + 1
		exit do
	ELSE
		read_row = read_row + 1
	END IF
	IF read_row = 18 THEN
		PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
		read_row = 7
	End if
LOOP until prog_type = "  "

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	script_end_procedure("Worker comments have been added to your notice. Please review the entire notice for accurarcy, and add additional comments as necessary.")
End if
