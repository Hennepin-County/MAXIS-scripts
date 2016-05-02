'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MA INMATE APPLICATION WCOM.vbs"
start_time = timer
'Required for statistical purposes==========================================================================================
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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



'--- DIALOGS-----------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dlg, 0, 0, 196, 85, "MA Inmate Application WCOM"
  EditBox 70, 15, 60, 15, case_number
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
  EditBox 75, 55, 60, 15, date_out
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

call MAXIS_case_number_finder(case_number)

'1st Dialog ---------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog case_number_dlg
	cancel_confirmation
	IF case_number = "" THEN err_msg = "Please enter a case number" & vbNewLine
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
CALL navigate_to_MAXIS_screen("STAT","HCRE")
'determining what row the member is on HCRE
hcre_row = 10
Do
	EMReadScreen Ref_number_hcre, 2, hcre_row, 24
	If Ref_number_hcre <> HH_member Then hcre_row = hcre_row + 1
	If hcre_row = 17 Then
		PF20 'shift pf7
		EMReadScreen hcre_message_check, 4, 24, 14    'checking to see if we hit the last page of HCRE
		IF hcre_message_check = "LAST" THEN script_end_procedure("Requested member was not found on HCRE. Please review case.") 'if the requested member isn't on HCRE end script
		hcre_row = 10
	END IF
Loop until Ref_number_hcre = HH_member
'Reading HCRE info for that member
EMReadScreen HCRE_coverage_date, 5, hcre_row, 64
'navigating to FACI to find which facility
CALL navigate_to_MAXIS_screen("STAT","FACI")
EMWriteScreen HH_member, 20, 76
EMWriteScreen "01", 20, 79
transmit
EMReadScreen FACI_total, 1, 2, 78
IF FACI_total = 0 THEN script_end_procedure("Correctional facility panel with an end date was not found for requested member. Please review case.")   'quitting if no FACI panels found.
If FACI_total <> 0 then 
    row = 14
    Do
		EMReadScreen faci_type, 2, 7, 43     'reading for facility type 68 (county correctional facility)
		IF faci_type = "68" THEN 
			EMReadscreen date_out, 10, row, 71    
			date_out = replace(date_out, " ", "/")
			If (left(date_out, 2) = left(HCRE_coverage_date, 2) AND right(date_out, 2) = right(HCRE_coverage_date, 2)) THEN     'the HCRE month matches the release month for this correctional facility
				EMReadScreen facility_name, 30, 6, 43
				facility_name = replace(facility_name, "_", "")
				Exit do
			ELSE
				row = row + 1
			END IF
				
			If row > 18 then
				EMReadScreen FACI_page, 1, 2, 73
				If FACI_page = FACI_total then       'if nothing is found stop script
					script_end_procedure("Correctional facility panel with an end date was not found for requested member. Please review case.")
				Else
					transmit
					row = 14
				End if
			End if
		ELSE
			EMReadScreen FACI_page, 1, 2, 73
			If FACI_page = FACI_total then        'if nothing is found stop script
				script_end_procedure("Correctional facility panel with an end date was not found for requested member. Please review case.")
			Else
				transmit
			END IF
		END IF
	Loop
End if


'2nd Dialog---------------------------------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog WCOM_dlg
	cancel_confirmation
	IF HH_member = "" THEN err_msg = err_msg & "Please enter your member number." & vbNewLine
	IF facility_name = "" THEN err_msg = err_msg & "Please enter your facility name." & vbNewLine
	IF isdate(date_out) = FALSE THEN err_msg = err_msg & "Please enter a valid date." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature" & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

call check_for_maxis(false)

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
	    EMSetCursor 03, 15
		CALL write_variable_in_SPEC_MEMO("MA begins " & date_out & ", the date you are released from the correctional facility.")
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
	call write_bullet_and_variable_in_CASE_NOTE("MA Start Date/Release Date", date_out)
	call write_bullet_and_variable_in_CASE_note("Facility", facility_name)
	call write_variable_in_CASE_note("* WCOM added to notice for member " & reference_number)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	script_end_procedure("Success! The WCOM/CASE NOTE have been added.")
END IF

script_end_procedure("")
