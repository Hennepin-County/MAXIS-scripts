'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - NEW SCRIPT TEMPLATE.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer

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


'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================

'DIALOGS FOR THE SCRIPT======================================================================================================

    '------Paste any dialogs needed in from the dialog editor here. Dialogs typically include MAXIS_case_number and worker_signature fields

'END DIALOGS=================================================================================================================

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
	err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
	Dialog sample_dialog                               'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
	IF ButtonPressed = cancel THEN StopScript          'If the user pushes cancel, stop the script
    
    'Handling for error messaging (in the case of mandatory fields or fields requiring a specific format)-----------------------------------
    'If a condition is met...          ...then the error message is itself, plus a new line, plus an error message...           ...Then add a comment explaining your reason it's mandatory.
	IF IsNumeric(MAXIS_case_number) = FALSE  THEN err_msg = err_msg & vbNewLine & "* You must type a valid numeric case number."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
	IF worker_signature = ""           THEN err_msg = err_msg & vbNewLine & "* You must sign your case note!"                  'worker_signature is usually also a mandatory field
    '<<Follow the above template to add more mandatory fields!!>>
    
    'If the error message isn't blank, it'll pop up a message telling you what to do!
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
'End dialog section-----------------------------------------------------------------------------------------------

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Now it navigates to a blank case note
start_a_blank_case_note

'...and enters a title (replace variables with your own content)...
CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")

'...some editboxes or droplistboxes (replace variables with your own content)...
CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )

'...checkbox responses (replace variables with your own content)...
If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )

'...and a worker signature. 
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")
