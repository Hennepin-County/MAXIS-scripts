'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - MNSURE MEMO.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 70                               'manual run time in seconds
STATS_denomination = "M"       'M is for each MEMBER
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MNsure_info_dialog, 0, 0, 196, 120, "MNsure Info Dialog"
  EditBox 60, 5, 70, 15, MAXIS_case_number
  DropListBox 110, 25, 75, 15, "denied"+chr(9)+"closed", how_case_ended
  EditBox 110, 45, 70, 15, denial_effective_date
  OptionGroup RadioGroup1
    RadioButton 20, 80, 35, 10, "WCOM", WCOM_check
    RadioButton 65, 80, 35, 10, "MEMO", MEMO_check
  EditBox 70, 100, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 80, 50, 15
    CancelButton 140, 100, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 100, 10, "Was case closed or denied?:"
  Text 5, 50, 100, 10, "Denial/closure effective date:"
  GroupBox 10, 70, 100, 25, "How are you sending this?"
  Text 5, 105, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog, checks for MAXIS or WCOM status.
Do
  Dialog MNsure_info_dialog
  cancel_confirmation
  If isdate(denial_effective_date) = False then MsgBox "You must put in a valid denial effective date (MM/DD/YYYY)."
Loop until isdate(denial_effective_date) = True

'checking for an active MAXIS session
check_for_maxis(FALSE)

CALL HH_member_custom_dialog(HH_member_array)

'For the WCOM option it needs to go to SPEC/WCOM. Otherwise it goes to MEMO.
If radiogroup1 = 0 then
	call navigate_to_MAXIS_screen("spec", "wcom")
	'Updates to show HC only memos
	EMWriteScreen "Y", 3, 74
	transmit

	FOR each HH_member in HH_member_array
		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"

		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen reference_number, 2, read_row, 62 'searching for selected HH members
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and reference_number = HH_member THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMReadScreen client_copy_check, 11, 1, 38
				If client_copy_check = "Client Copy" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
				'Sends the home key to get to the top of the memo.
				EMSendKey "<home>"

				'Enters different text for denials vs closures. This adds the different text to the first line
				If how_case_ended = "denied" then EMSendKey "Your application was denied "
				If how_case_ended = "closed" then EMSendKey "Your case was closed "

				'Now it sends the rest of the memo, saves the memo and exits the memo screen
				EMSendKey "effective " & denial_effective_date & "." & "<newline>" & "<newline>" & "You may be able to purchase medical insurance through MNsure. If your family is under an income limit you may get financial help to purchase insurance. You can apply online at www.mnsure.org. If you have questions or need help to apply you can call the MNsure Call Center at 1-855-366-7873."
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
		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
	NEXT
	If WCOM_count = 0 THEN
		MSGbox "No Waiting HC elig results were found in this month for this HH member."
		Stopscript
	END IF
Else
	'Navigating to SPEC/MEMO
	call navigate_to_MAXIS_screen("SPEC", "MEMO")
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMWriteScreen "x", 5, 10
	transmit
	'Sends the home key to get to the top of the memo.
	EMSendKey "<home>"

	'Enters different text for denials vs closures. This adds the different text to the first line
	If how_case_ended = "denied" then EMSendKey "Your application was denied "
	If how_case_ended = "closed" then EMSendKey "Your case was closed "

	'Now it sends the rest of the memo, saves the memo and exits the memo screen
	EMSendKey "effective " & denial_effective_date & "." & "<newline>" & "<newline>" & "You may be able to purchase medical insurance through MNsure. If your family is under an income limit you may get financial help to purchase insurance. You can apply online at www.mnsure.org. If you have questions or need help to apply you can call the MNsure Call Center at 1-855-366-7873."
	PF4
	PF3
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
End if


'Enters case note
start_a_blank_CASE_NOTE
If radiogroup1 = 0 then EMSendKey "Added MNsure info to client notice via WCOM. -" & worker_signature
If radiogroup1 = 1 then EMSendKey "Sent client MNsure info via MEMO. -" & worker_signature

STATS_counter = STATS_counter - 1                      'subtracts one instance to the stats counter
script_end_procedure("")
