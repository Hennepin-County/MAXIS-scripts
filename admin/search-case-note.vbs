'STATS GATHERING=============================================================================================================
name_of_script = "ADMIN - Search CASE NOTE.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'FUNCTIONS FOR THE SCRIPT====================================================================================================

function go_to_top_of_notes()
	Do
		PF7
		EMReadScreen top_of_notes_check, 10, 24, 14
	Loop until top_of_notes_check = "FIRST PAGE"
end function

'END FUNCCCTIONS=============================================================================================================

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""
CALL check_for_MAXIS(True)

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Initial Dialog - Case number
Dialog1 = ""                                        'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 166, 130, "Search in CASE:NOTE"
  EditBox 60, 35, 45, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 65, 90, 95, 15, "Script Instructions", script_instructions_btn
    OkButton 55, 110, 50, 15
    CancelButton 110, 110, 50, 15
  Text 5, 10, 155, 20, "This script can search for words or phrases in CASE:NOTE for a specific case."
  Text 5, 40, 50, 10, "Case Number:"
  Text 5, 65, 155, 20, "We can only search one case at a time, ensure you have the correct CASE NUMBER here."
EndDialog


'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
        If ButtonPressed = script_instructions_btn Then             'Pulling up the instructions if the instruction button was pressed.
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx"
            err_msg = "LOOP"
        Else                                                        'If the instructions button was NOT pressed, we want to display the error message if it exists.
		    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        End If
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

Dim search_array()
ReDim search_array(0)
search_item_found = False

'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
	Original_search = search_text
	If search_array(0) = "" OR search_item_found = True Then
		BeginDialog Dialog1, 0, 0, 316, 90, "Search " & MAXIS_case_number & " CASE:NOTEs"
		  EditBox 50, 10, 260, 15, search_text
		  DropListBox 50, 30, 60, 45, "AND"+chr(9)+"OR", search_type
		  ButtonGroup ButtonPressed
		    PushButton 250, 30, 60, 15, "Search", search_button
		    If search_array(0) <> "" Then PushButton 190, 70, 70, 15, "Find Next - PF3", find_next_button
		    PushButton 260, 75, 50, 10, "Done", done_button
		  Text 5, 15, 45, 10, "Search Text:"
		  Text 5, 35, 45, 10, "Search Type:"
		  Text 5, 55, 305, 10, read_the_line
		EndDialog

		err_msg = ""                                       'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
		Dialog Dialog1                               'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
		If ButtonPressed = done_button Then ButtonPressed = 0
		cancel_without_confirmation
		If ButtonPressed = -1 Then
			If Original_search = search_text Then ButtonPressed = find_next_button
			If Original_search <> search_text Then ButtonPressed = search_button
		End If
	End If

	If ButtonPressed = search_button Then
		EMReadScreen check_we_are_in_case_note, 17, 2, 33
		If check_we_are_in_case_note <> "Case Notes (NOTE)" then call navigate_to_MAXIS_screen("CASE", "NOTE")

		go_to_top_of_notes

		case_to_read_row = 5
		line_to_read_row = 4

		ReDim search_array(0)
		search_array(0) = ""
		If search_type = "AND" Then
			search_array(0) = UCase(search_text)
		End If
		If search_type = "OR" Then
			temp_array = split(search_text, " ")
			search_item_counter = 0
			For each search_item in temp_array
				ReDim Preserve search_array(search_item_counter)
				search_array(search_item_counter) = UCase(search_item)
				search_item_counter = search_item_counter + 1
			Next
		End If
	End If

	on_note_menu = False
	EMReadScreen are_we_on_note_menu, 10, 4, 25
	If are_we_on_note_menu = "First line" Then on_note_menu = True

	If on_note_menu = True Then
		EMWriteScreen "X", case_to_read_row, 3
		transmit
		line_to_read_row = 4
	End If

	EMReadScreen read_the_line, 78, line_to_read_row, 3
	read_the_line = UCase(read_the_line)

	For each search_item in search_array
		If Instr(read_the_line, search_item) <> 0 Then
			search_item_found = True
			EMSetCursor line_to_read_row, Instr(read_the_line, search_item) + 2
			line_to_read_row = line_to_read_row + 1
			Exit For
		End If
	Next
	If line_to_read_row = 18 Then
		PF8
		line_to_read_row = 4
		EmReadScreen end_of_note, 9, 24, 14
		If end_of_note = "LAST PAGE" Then
			PF3
			case_to_read_row = case_to_read_row + 1
		End If
	End If

LOOP
'End dialog section-----------------------------------------------------------------------------------------------

' 'Checks Maxis for password prompt
' CALL check_for_MAXIS(True)
'
' 'Now it navigates to a blank case note
' start_a_blank_case_note
'
' '...and enters a title (replace variables with your own content)...
' CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")
'
' '...some editboxes or droplistboxes (replace variables with your own content)...
' CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
' CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )
'
' '...checkbox responses (replace variables with your own content)...
' If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
' If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )
'
' '...and a worker signature.
' CALL write_variable_in_case_note("---")
' CALL write_variable_in_case_note(worker_signature)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")
