'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - ELIG RESULTS TO WORD.vbs"
start_time = timer
STATS_counter = 0
STATS_manualtime = 60          'manual run time in seconds
STATS_denomination = "C"        'C is for each case

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
call changelog_update("06/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function copy_elig_to_array(output_array)
	output_array = "" 'resetting array
	For row = 2 to 21
		EMReadScreen elig_line, 80, row, 1
		output_array = output_array & elig_line & "UUDDLRLRBA"
	Next
	' MsgBox output_array
	output_array = split(output_array, "UUDDLRLRBA")
End function

function insert_page_break_after_two_panels(screen_on_page)
	'Determines if the Word doc needs a new page
	'screen_on_page - This is a running counter that is updated in this function
	screen_on_page = screen_on_page + 1
	If screen_on_page = 3 then												'if we are at 2, we need to insert a page breakk and reset the counter
		screen_on_page = 1
		objSelection.InsertBreak(7)
	ElseIf screen_on_page = 2 Then
		objSelection.TypeText vbCr & vbCr & vbCr & vbCr & vbCr
	End if
	STATS_counter = STATS_counter + 1											'also using this to increment the stats counter since we do this with every panel we read.'
end function
screen_on_page = 0

'THE SCRIPT ================================================================================================================
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 231, 195, "Case Number to Read ELIG Results"
  EditBox 65, 65, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 115, 170, 50, 15
    CancelButton 170, 170, 50, 15
  Text 15, 15, 175, 20, "This script works by pulling all of the information from a specific version of ELIG into word."
  Text 15, 40, 195, 20, "Enter the Case Number and Navigate in MAXIS now to the start of the ELIG version that you would like copied."
  Text 15, 70, 45, 10, "Case Number"
  GroupBox 10, 90, 205, 40, "NAVIGATE IN MAXIS NOW TO THE ELIG VERSION TO COPY"
  Text 25, 105, 165, 10, "- ELIG can be for any program except HC."
  Text 25, 115, 165, 10, "- Version of ELIG can be approved or unapproved."
  Text 15, 140, 180, 20, "When the script continues, it will look for the first page of ELIG Results. If not found, the dialog will reappear."
EndDialog

Do
	Do
		Do
			err_msg = ""

			dialog Dialog1
			cancel_without_confirmation

			Call validate_MAXIS_case_number(err_msg, "*")

		Loop until err_msg = ""
		elig_results_program_found = ""
		EMReadScreen MX_line_3, 78, 3, 2
		If InStr(MX_line_3, "DWPR") Then elig_results_program_found = "DWP"
		If InStr(MX_line_3, "MFPR") Then elig_results_program_found = "MFIP"
		If InStr(MX_line_3, "MFSC") Then elig_results_program_found = "MFIP"
		If InStr(MX_line_3, "MSPR") Then elig_results_program_found = "MSA"
		If InStr(MX_line_3, "GAPR") Then elig_results_program_found = "GA"
		If InStr(MX_line_3, "CAPR") Then elig_results_program_found = "Cash Denial"
		If InStr(MX_line_3, "GRPR") Then elig_results_program_found = "GRH"
		If InStr(MX_line_3, "FCSM") Then elig_results_program_found = "IV-E"
		If InStr(MX_line_3, "EMPR") Then elig_results_program_found = "EMER"
		If InStr(MX_line_3, "FSPR") Then elig_results_program_found = "SNAP"
		If elig_results_program_found = "" Then MsgBox "MAXIS must be at the first page of Eligibility Results for this script to run." & vbCr & vbCr & "The dialog will return." & vbCr & vbCr & "NAVIGATE TO ELIG RESULTS WHILE THE DIALOG IS UP."

		EMReadScreen Look_at_MAXIS_word, 12, 1, 36
		Look_at_MAXIS_word = trim(Look_at_MAXIS_word)
		If Look_at_MAXIS_word <> "MAXIS" Then MsgBox "You do not appear to be in MAXIS, navigate to MAXIS when the dialog reappears."

		EMReadScreen version_number, 3, 2, 11
	Loop until elig_results_program_found <> "" and Look_at_MAXIS_word = "MAXIS"
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Now we make sure the check for password transmit wasn't a problem.
'Need to navigate to the first page based on program. Each program has different coordinates

If elig_results_program_found = "DWP" Then Call write_value_and_transmit("DWPR", 20, 71)
If elig_results_program_found = "MFIP" Then Call write_value_and_transmit("MFSC", 20, 71)
If elig_results_program_found = "MSA" Then Call write_value_and_transmit("MSPR", 20, 71)
If elig_results_program_found = "GA" Then Call write_value_and_transmit("GAPR", 20, 70)
If elig_results_program_found = "Cash Denial" Then Call write_value_and_transmit("CAPR", 19, 70)
If elig_results_program_found = "GRH" Then Call write_value_and_transmit("GRPR", 20, 70)
'IVE does not need to navigate back because there is only one ELIG page
If elig_results_program_found = "EMER" Then Call write_value_and_transmit("EMPR", 19, 70)
If elig_results_program_found = "SNAP" Then Call write_value_and_transmit("FSPR", 19, 70)

'confirm we are at the right version
EMReadScreen check_version_number, 3, 2, 11
If check_version_number <> version_number Then
	write_version_number = trim(version_number)
	write_version_number = right("00"&write_version_number, 2)

	row = 1
	col = 1
	EMSearch "Command:", row, col
	Call write_value_and_transmit(write_version_number, row, col +17)
End If

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 50
objSelection.PageSetup.BottomMargin = 50
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "10"

'Read each line of the results.
elig_line_array = ""
Do
	Call copy_elig_to_array(elig_line_array)

	Call insert_page_break_after_two_panels(screen_on_page)

	'Adds current screen to Word doc
	For each line in elig_line_array
		objSelection.TypeText line & Chr(11)
	Next

	'POP-UPs not handled at this time in the Word Document

	transmit

	last_elig_screen = False
	EMReadScreen worker_message, 78, 24, 2
	worker_message = trim(worker_message)
	If worker_message = "PROVIDE A COMMAND OR PF-KEY TO CONTINUE" Then last_elig_screen = True
	If worker_message = "** PLEASE PROVIDE A COMMAND OR PF-KEY TO CONTINUE" Then last_elig_screen = True
	If worker_message = "PLEASE PROVIDE A COMMAND OR PF-KEY TO CONTINUE" Then last_elig_screen = True
Loop Until last_elig_screen = True

'Document 'meta-data' information
objSelection.TypeParagraph()
objSelection.TypeText "===============================================================================" & vbCr
objSelection.TypeText " Eligibility Information captured on: " & date & " (date Word Document created)" & vbCr
objSelection.TypeText "==============================================================================="

Call script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/12/2024
'--Tab orders reviewed & confirmed----------------------------------------------03/12/2024
'--Mandatory fields all present & Reviewed--------------------------------------03/12/2024
'--All variables in dialog match mandatory fields-------------------------------03/12/2024
'Review dialog names for content and content fit in dialog----------------------03/12/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function:
'    confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/12/2024					Did not use check_for_MAXIS function because of transmit - there is specialized support here
'--MAXIS_background_check reviewed (if applicable)------------------------------03/12/2024
'--PRIV Case handling reviewed -------------------------------------------------N/A							Not needed because the script only works if in ELIG - which is not possible if PRIV
'--Out-of-County handling reviewed----------------------------------------------N/A							There is no out of county restriction.
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/12/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------03/12/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------03/12/2024
'--Incrementors reviewed (if necessary)-----------------------------------------03/12/2024
'--Denomination reviewed -------------------------------------------------------03/12/2024
'--Script name reviewed---------------------------------------------------------03/12/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------03/12/2024
'--comment Code-----------------------------------------------------------------03/12/2024
'--Update Changelog for release/update------------------------------------------03/12/2024
'--Remove testing message boxes-------------------------------------------------03/12/2024
'--Remove testing code/unnecessary code-----------------------------------------03/12/2024
'--Review/update SharePoint instructions----------------------------------------03/13/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------03/13/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------03/13/2024