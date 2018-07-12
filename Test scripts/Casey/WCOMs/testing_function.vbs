'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SELECT WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================
'run_locally = TRUE
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
FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/13/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function write_thing_in_SPEC_MEMO(variable)
'--- This function writes a variable in SPEC/MEMO
'~~~~~ variable: information to be entered into SPEC/MEMO
'===== Keywords: MAXIS, SPEC, MEMO
	EMGetCursor memo_row, memo_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	memo_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page
	Do
		EMReadScreen line_test, 60, memo_row, memo_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
        line_test = trim(line_test)
        'MsgBox line_test
        If line_test <> "" OR memo_row >= 18 Then
            memo_row = memo_row + 1

            'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
            If memo_row >= 18 then
                PF8
                memo_row = 3					'Resets this variable to 3
            End if
        End If

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then script_end_procedure("The WCOM/MEMO area is already full and no additional informaion can be added. This script should be run prior to adding manual wording.")

	Loop until line_test = ""

	'Each word becomes its own member of the array called variable_array.
	variable_array = split(variable, " ")

	For each word in variable_array
		'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
		If len(word) + memo_col > 74 then
			memo_row = memo_row + 1
			memo_col = 15
		End if

		'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
		If memo_row >= 18 then
			PF8
			memo_row = 3					'Resets this variable to 3
		End if

        EMReadScreen page_full_check, 12, 24, 2
        'MsgBox page_full_check
        If page_full_check = "END OF INPUT" Then
            PF10
            end_msg = "The WCOM/MEMO area is already full and no additional informaion can be added. The wording that was not added and the script ended on is:" & vbNewLine & vbNewLine & variable & vbNewLine & vbNewLine & "**This script should be run prior to adding manual wording.**"
            script_end_procedure(end_msg)
        End If
		'Writes the word and a space using EMWriteScreen
		EMWriteScreen word & " ", memo_row, memo_col

		'Increases memo_col the length of the word + 1 (for the space)
		memo_col = memo_col + (len(word) + 1)
	Next

	'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand).
	EMSetCursor memo_row + 1, 15
end function


'THE SCRIPT
EMConnect ""            'Connect to BlueZone

msg_line = "You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your team if you have questions."

'Open the Notice
EMWriteScreen "X", 8, 13
transmit

PF9     'Put in to edit mode - the worker comment input screen
EMSetCursor 03, 15

CALL write_thing_in_SPEC_MEMO(msg_line)
