'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - MA INMATE APPLICATION WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
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
call changelog_update("04/04/2017", "Added handling for multiple recipient changes to SPEC/WCOM", "David Courtright, St Louis County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'--- DIALOGS-----------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dlg, 0, 0, 196, 85, "MA Inmate Application WCOM"
  EditBox 70, 15, 60, 15, MAXIS_case_number
  EditBox 70, 35, 30, 15, approval_month
  EditBox 160, 35, 30, 15, approval_year
  ButtonGroup ButtonPressed
    OkButton 45, 60, 50, 15
    CancelButton 100, 60, 50, 15
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
EndDialog

BeginDialog WCOM_dlg, 0, 0, 146, 120, "MA Inmate Application WCOM"
  EditBox 75, 15, 60, 15, HH_member
  EditBox 75, 35, 60, 15, facility_name
  EditBox 75, 55, 60, 15, most_recent_release_date
  EditBox 75, 75, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 15, 95, 50, 15
    CancelButton 80, 95, 50, 15
  Text 10, 40, 60, 10, "Facility Name:"
  Text 10, 60, 60, 10, "MA Start Date:"
  Text 10, 20, 60, 10, "Member Number:"
  Text 10, 80, 60, 10, "Worker Signature:"
EndDialog


'--------------------------------------------------------------------------------------------------------------------------------

'--- The script -----------------------------------------------------------------------------------------------------------------

EMConnect ""

call MAXIS_case_number_finder(MAXIS_case_number)

'1st Dialog ---------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog case_number_dlg
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = "Please enter a case number" & vbNewLine
	IF len(approval_month) <> 2 THEN err_msg = err_msg & "Please enter your month in MM format." & vbNewLine
	IF len(approval_year) <> 2 THEN err_msg = err_msg & "Please enter your year in YY format." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

call check_for_maxis(false)

'Creating HH member array-------------------------------------------------------------------------------------------------------------
DO							'Loops until worker selects only one HH member. At this time the script only handles one HH member due to grammar issues involving multiple members with different postponed WREG verifs.
	Msgbox "Select the HH member that has had the inmate application approved. If you have multiple HH members please process manually at this time."
	CALL HH_member_custom_dialog(HH_member_array)
	array_length = Ubound(HH_member_array)
LOOP until array_length = 0

HH_member = HH_member_array(0)

call check_for_maxis(false)

'Gathering/formatting variables---------------------------------------------------------------------------------------------------------------------
back_to_self
'navigating to FACI to find which facility
CALL navigate_to_MAXIS_screen("STAT","FACI")
EMWriteScreen HH_member, 20, 76
EMWriteScreen "01", 20, 79
transmit
EMReadScreen FACI_total, 1, 2, 78
IF FACI_total = 0 THEN script_end_procedure("Correctional facility panel with an end date was not found for requested member. Please review case.")   'quitting if no FACI panels found.
If FACI_total <> 0 then
DO
	row = 14
	EMReadScreen FACI_current, 1, 2, 73
    Do
		EMReadScreen faci_type, 2, 7, 43     'reading for facility type 68 (county correctional facility) 69 (non county correctional facility)
		IF faci_type = "68" or faci_type = "69" THEN
			EMReadscreen date_out, 10, row, 71
			If date_out = "__ __ ____" AND row = 14 THEN Exit Do										'stopping as if the first row read doesn't have an end date then it cannot be compared.
			If date_out <> "__ __ ____" or date_out = "          " THEN 							'finding the most recent date out. Per MAXIS this will always be the one on the bottom
				row = row + 1																		'if it finds anything other than a blank field it goes to the next row.
			ELSE
				EMReadscreen date_out, 10, row - 1, 71											'once it finds a blank row it looks are the row above it (the most recent out date for that panel)
				date_out = replace(date_out, " ", "/")
				IF date_to_compare = "" THEN 													'it now sets a date to compare by if it's the first time through the loop that bar is set here.
					date_to_compare = date_out
					previous_date_diff = datediff("d", date_to_compare, date_out)
				END IF
				IF previous_date_diff =< datediff("d", date_to_compare, date_out) Then			'here it actually compares the overall most recent date with the most recent date found on current panel
					previous_date_diff = datediff("d", date_to_compare, date_out)				'resetting the most recent date to compare with if a most recent date is found
					EMReadScreen facility_name, 30, 6, 43										'defining the facility if the current most recent date is found
					most_recent_release_date = cstr(date_out)									'converting as dialogs won't display dates sometimes
					facility_name = replace(facility_name, "_", "")								'cleaning up the facility name
				END IF
				exit do
			END If
		ELSE
			Exit Do
		END IF
	Loop until row = 19											'looping until there are no more date outs to be read on that panel.
	Transmit
Loop until FACI_current = FACI_total														'looping until you've checked all of the panels available.
END IF

IF facility_name = "" THEN script_end_procedure("The script was unable to find a FACI panel with 68 or 69 and an end date. Please review FACI panel.")

'2nd Dialog---------------------------------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog WCOM_dlg
	cancel_confirmation
	IF HH_member = "" THEN err_msg = err_msg & "Please enter your member number." & vbNewLine
	IF facility_name = "" THEN err_msg = err_msg & "Please enter your facility name." & vbNewLine
	IF isdate(most_recent_release_date) = FALSE THEN err_msg = err_msg & "Please enter a valid date." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature" & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""
'This section will check for whether forms go to AREP and SWKR
call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
EMReadscreen forms_to_arep, 1, 10, 45
call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
EMReadscreen forms_to_swkr, 1, 15, 63

'WCOM PIECE---------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("spec", "wcom")

EMWriteScreen approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
EMWriteScreen "Y", 3, 74 'selects HC only
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	waiting_check = ""
	EMReadscreen reference_number, 2, read_row, 62
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and reference_number = HH_member THEN 'checking program type and if it's been printed
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
		CALL write_variable_in_SPEC_MEMO("MA begins " & most_recent_release_date & ", the date you are released from the correctional facility.")
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
LOOP until reference_number = "  "

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting HC elig results were found in this month for this HH member. Please review ELIG results.")
ELSE 					'If a waiting FS notice is found
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM Regarding Inmate Application Added---")
	call write_bullet_and_variable_in_CASE_NOTE("MA Start Date/Release Date", most_recent_release_date)
	call write_bullet_and_variable_in_CASE_note("Facility", facility_name)
	call write_variable_in_CASE_note("* WCOM added to notice for member " & reference_number)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	script_end_procedure("Success! The WCOM/CASE NOTE have been added.")
END IF

script_end_procedure("")
