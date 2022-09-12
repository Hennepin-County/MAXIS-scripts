'STATS GATHERING=============================================================================================================
name_of_script = "UTILITIES - COPY CASE NOTE TO WORD.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 0                     	'sets the stats counter at one
STATS_manualtime = 4                	'manual run time in seconds
STATS_denomination = "I"       		'I is for each ITEM
'END OF stats block==========================================================================================================

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

''CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("08/04/2021", "Updated to work. Previously the script errored with every run", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS ==================================================================================================================
function read_case_note_for_word(output_array)
	For start_rows = 1 to 3
		EMReadScreen reading_line, 80, start_rows, 1
		output_array = output_array & reading_line & "UUDDLRLRBA"
	Next

	row = 4
	Do
		EMReadScreen reading_line, 80, row, 1
		' MsgBox "row - " & row & vbCr & "reading_line - " & reading_line
		output_array = output_array & reading_line & "UUDDLRLRBA"

		row = row + 1
		If row = 18 Then
			EMReadScreen more_note, 7, 18, 3
			If more_note = "More: +" Then
				PF8
				row = 4
			End If
		End If
	Loop until more_note <> "More: +" and more_note <> ""

	For end_rows = 18 to 24
		EMReadScreen reading_line, 80, end_rows, 1
		output_array = output_array & reading_line & "UUDDLRLRBA"
	Next

	output_array = split(output_array, "UUDDLRLRBA")
end function

function insert_page_break_after_two_panels(screen_on_page)
	'Determines if the Word doc needs a new page
	'screen_on_page - This is a running counter that is updated in this function
	If screen_on_page = "" or screen_on_page = 1 then							'if we are at 1, we need to add some spaces and increment the counter'
		screen_on_page = 2
		objSelection.TypeText vbCr & vbCr
	Elseif screen_on_page = 2 then												'if we are at 2, we need to insert a page breakk and reset the counter
		screen_on_page = 1
		objSelection.InsertBreak(7)
	End if
	STATS_counter = STATS_counter + 1											'also using this to increment the stats counter since we do this with every panel we read.'
end function
'END FUNCTIONS BLOCK ========================================================================================================

'THE SCRIPT==================================================================================================================
screen_on_page = 1

'Connects to BlueZone
EMConnect ""
Call check_for_MAXIS(True)			'ensuring we are logged in to MAXIS.

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

'This is a dialog just so that we can naviagate while it is up - there are no inputs.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 85, "Ensure in CASE NOTE"
  ButtonGroup ButtonPressed
    OkButton 160, 65, 50, 15
    CancelButton 210, 65, 50, 15
  Text 10, 10, 255, 10, "This script will pull the details of a single CASE/NOTE into a Word Documnet."
  Text 10, 25, 245, 10, "At this time, ensure you are in a CASE/NOTE in MAXIS that has been saved."
  GroupBox 10, 40, 115, 40, "NAVIGATE TO CORRECT NOTE"
  Text 20, 55, 100, 15, "Manually navigate to the correct CASE/NOTE in MAXIS now."
EndDialog

DO
	Dialog Dialog1                                'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
	cancel_confirmation

	Call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = False

'making sure we are actually in a CASE/NOTE'
EMReadScreen check_in_case_note, 10, 2, 33
EMReadScreen case_note_mode, 7, 20, 3

If check_in_case_note <> "Case Notes" or case_note_mode <> "Mode: D" Then script_end_procedure("The script was not in a CASE/NOTE. Navigate into a CASE/NOTE and run again if you have a NOTE you would like in WORD.")

'going to the top of the note.
Do
	PF7
	EMReadScreen more_note, 9, 18, 3
Loop Until more_note <> "More: +/-" and more_note <> "More:   -"
more_note = ""

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 30
objSelection.PageSetup.BottomMargin = 30
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "10"

'grabbing the CASE Number for stats
EMReadScreen MAXIS_case_number, 8, 20, 38
MAXIS_case_number = trim(MAXIS_case_number)

call read_case_note_for_word(screentest)		'reads all the lines of the current note

'Adds current screen to Word doc
For each line in screentest
	objSelection.TypeText line & Chr(11)
	STATS_counter = STATS_counter + 1											'also using this to increment the stats counter since we do this with everyrow we write'
Next

script_end_procedure("Success! Word Document created and opened with CASE/NOTE text.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/12/2022
'--Tab orders reviewed & confirmed----------------------------------------------09/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------N/A
'--All variables in dialog match mandatory fields-------------------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/12/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------09/12/2022
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/12/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/12/2022
'--Incrementors reviewed (if necessary)-----------------------------------------09/12/2022
'--Denomination reviewed -------------------------------------------------------09/12/2022
'--Script name reviewed---------------------------------------------------------09/12/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------09/12/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/12/2022
'--comment Code-----------------------------------------------------------------09/12/2022
'--Update Changelog for release/update------------------------------------------09/12/2022
'--Remove testing message boxes-------------------------------------------------N/A
'--Remove testing code/unnecessary code-----------------------------------------N/A
'--Review/update SharePoint instructions----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
