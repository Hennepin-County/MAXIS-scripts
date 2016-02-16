'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - ABAWD WITH CHILD IN HH WCOM.vbs"
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
STATS_manualtime = 60                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'--- DIALOGS-----------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dlg, 0, 0, 196, 100, "Adding ABAWD Adult WCOM"
  EditBox 70, 15, 60, 15, case_number
  EditBox 70, 35, 30, 15, approval_month
  EditBox 160, 35, 30, 15, approval_year
  EditBox 95, 55, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 80, 50, 15
    CancelButton 100, 80, 50, 15
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
  Text 25, 60, 70, 10, "Worker signature: "
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
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

call check_for_maxis(false)

'Creating HH member array-------------------------------------------------------------------------------------------------------------
Msgbox "Select the SNAP ABAWD member(s) that is exempt due to a child under 18."
CALL HH_member_custom_dialog(HH_member_array)


call check_for_maxis(false)

'Gathering/formatting variables---------------------------------------------------------------------------------------------------------------------
back_to_self
EMWriteScreen approval_month, 20, 43
EMWriteScreen approval_year, 20, 46
CALL check_for_maxis(false)
FOR each HH_member in HH_member_array
	CALL navigate_to_MAXIS_screen("STAT", "MEMB")  'grabbing client's name
	EMWriteScreen HH_member, 20, 76
	Transmit
	EMReadScreen First_name, 12, 6, 63
	EMReadScreen Last_name, 25, 6, 30
	EMReadScreen Middle_initial, 1, 6, 79
	client_name = client_name & replace(First_name, "_", "") & " " & replace(Middle_initial, "_", "") & " " & replace(Last_name, "_", "") & ", "
NEXT

'WCOM PIECE---------------------------------------------------------------------------------------------------------------------
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
	    EMSetCursor 03, 15
    	CALL write_variable_in_SPEC_MEMO(client_name & " is(are) exempt from Able Bodied Adult Without Dependents(ABAWD) work requirements due to a child(ren) under the age of 18 in the SNAP unit. ")
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

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	script_end_procedure("Success! The WCOM/CASE NOTE/TIKL have been added.")
END IF

script_end_procedure("")
