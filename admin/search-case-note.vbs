'STATS GATHERING=============================================================================================================
name_of_script = "UTILITIES - Search CASE NOTE.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 0               'sets the stats counter at one
STATS_manualtime = 5            'manual run time in seconds
STATS_denomination = "I"        'C is for each case
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
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/17/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
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
CALL check_for_MAXIS(True)			'Making sure we are signed in - ends the script run if we are not

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
            run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ADMIN/ADMIN%20-%20SEARCH%20CASE%20NOTE.docx"
            err_msg = "LOOP"
        Else                                                        'If the instructions button was NOT pressed, we want to display the error message if it exists.
		    IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        End If
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'declaring some defaults
Dim search_array()
ReDim search_array(0)
search_item_found = False
end_of_notes = False

'Here we loop a lot. The rest of the sript is this loop.
'Basically we are using this dialog to enter the search parameter and display the information found in the search.
DO
	Original_search = search_text												'defining what we were previously searching so we know if it is changed/new
	If search_array(0) = "" OR search_item_found = True OR end_of_notes = True Then						'we will not display the dialog on every loop - only if we need to display a search or if there is no parameter
		Do
			Dialog1 = ""                                        'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 316, 120, "Search " & MAXIS_case_number & " CASE:NOTEs"
			  EditBox 50, 10, 260, 15, search_text
			  DropListBox 50, 30, 60, 45, "OR"+chr(9)+"AND", search_type
			  If search_item_found = True Then									'only show the search results if there is a result of the search
				  Text 5, 50, 35, 10, "Date: "
				  Text 40, 50, 100, 10, case_note_date
				  Text 5, 60, 35, 10, "Header: "
				  Text 40, 60, 265, 10, case_note_header
				  Text 5, 75, 305, 20, case_note_line_display
			  End If
			  If end_of_notes = True Then										'only display the information about the notes end being reached if that happened
			  	Text 5, 50, 305, 30, "All current notes have been searched for " & join(search_array, ", ")
			  End If
			  ButtonGroup ButtonPressed
			    PushButton 250, 30, 60, 15, "Search", search_button
			    If search_array(0) <> "" AND end_of_notes = False Then PushButton 190, 100, 70, 15, "Find Next - PF3", find_next_button
			    PushButton 260, 105, 50, 10, "Done", done_button
			  Text 5, 15, 45, 10, "Search Text:"
			  Text 5, 35, 45, 10, "Search Type:"
			EndDialog

			Dialog Dialog1                               						'The Dialog command shows the dialog.
			If ButtonPressed = done_button Then ButtonPressed = 0
			cancel_without_confirmation
			search_item_found = False		'resetting this after the display
			If ButtonPressed = -1 Then		'defaulting what happens if the user presses the 'enter' button
				If Original_search = search_text Then ButtonPressed = find_next_button
				If Original_search <> search_text Then ButtonPressed = search_button
			End If
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = False
	End If

	'if there is a new search to complete, this code gets us back to starting from the beginning
	If ButtonPressed = search_button Then
		EMReadScreen check_we_are_in_case_note, 17, 2, 33						'making sure we are in CASE:NOTES'
		If check_we_are_in_case_note <> "Case Notes (NOTE)" then call navigate_to_MAXIS_screen("CASE", "NOTE")
		on_note_menu = False													'making sure we are on the main menu of CASE:NOTES'
		EMReadScreen are_we_on_note_menu, 10, 4, 25
		If are_we_on_note_menu = "First line" Then on_note_menu = True
		If on_note_menu = False Then PF3										'backing out of the note if we are in it'

		go_to_top_of_notes														'make sure we haven't paged down in CASE::NOTES
		end_of_notes = False

		case_to_read_row = 5		'defaulting which row to start at
		line_to_read_row = 4

		ReDim search_array(0)		''resetting the search array
		search_array(0) = ""
		If search_type = "AND" Then												'creating an array with a single item of the ALL of the search critera
			search_array(0) = UCase(search_text)
		End If
		If search_type = "OR" Then												'creating an array of all of the words in the search field
			temp_array = split(search_text, " ")
			search_item_counter = 0
			For each search_item in temp_array
				ReDim Preserve search_array(search_item_counter)
				search_array(search_item_counter) = UCase(search_item)
				search_item_counter = search_item_counter + 1
			Next
		End If
		ButtonPressed = find_next_button										'making sure the search button is not the pressed button on the next loop
	End If

	on_note_menu = False														'checking to see if we are on the menu of the CASE:NOTES
	EMReadScreen are_we_on_note_menu, 10, 4, 25
	If are_we_on_note_menu = "First line" Then on_note_menu = True

	If on_note_menu = True Then													'if we are on the menu we need to open the note
		EMWriteScreen "X", case_to_read_row, 3
		EMReadScreen case_note_date, 8, case_to_read_row, 6						'grabbing the date for display
		If case_note_date = "        " Then end_of_notes = True
		transmit
		EMReadScreen case_note_header, 78, 4, 3									'grabbing the ccase note header for display
		line_to_read_row = 4
	End If

	STATS_counter = STATS_counter + 1											'incrementing the counter for each line read'
	EMReadScreen read_the_line, 78, line_to_read_row, 3							'reading a line of the note
	case_note_line_display = read_the_line										'creeating a display variable of the line
	read_the_line = UCase(read_the_line)										'ucase the line so we are comparing capitals to capitals'

	For each search_item in search_array										'checking each item in the search parameters to see if it is somewhere in the line that was read.
		If Instr(read_the_line, search_item) <> 0 Then							'if the instring was found
			search_item_found = True											'identify in the boolean that it was found so the dialog displays
			EMSetCursor line_to_read_row, Instr(read_the_line, search_item) + 2	'set the cursor on the found parameter
			Exit For
		End If
	Next
	line_to_read_row = line_to_read_row + 1										'going to the next line in the case:note
	If line_to_read_row = 18 Then												'if we are at line 18 - this page is at the end
		PF8																		'go to the next page of the case:note
		line_to_read_row = 4													'reset the line to read to the top of the case:note
		EMReadScreen end_of_note, 9, 24, 14										'read to see if we go a message that we are at the last page - this only displays when you try to PF8 from the last page, NOT when you arrive at the last page
		If end_of_note = "LAST PAGE" Then										'if the message indicates we are already at the last page
			PF3																	'leave the current case:note
			case_to_read_row = case_to_read_row + 1								'go to the next line of notes
			If case_to_read_row = 19 Then										'if we are at row 19 - this page of case:notes is at the end
				PF8																'go to the next row
				case_to_read_row = 5											'reset the row to start opening notes'
				EMReadScreen end_of_all_notes, 9, 24, 14						'read the message that will display if we PF8 from the last page of notes
				If end_of_all_notes = "LAST PAGE" Then							'if it displays last page, there are no more notes to read.
					end_of_notes = True
				End If
			End If
		End If
	End If
LOOP

script_end_procedure("")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/17/2021
'--Tab orders reviewed & confirmed----------------------------------------------11/17/2021
'--Mandatory fields all present & Reviewed--------------------------------------N/A
'--All variables in dialog match mandatory fields-------------------------------N/A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------11/17/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------11/17/2021
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/17/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/17/2021
'--Incrementors reviewed (if necessary)-----------------------------------------11/17/2021
'--Denomination reviewed -------------------------------------------------------11/17/2021
'--Script name reviewed---------------------------------------------------------11/17/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------11/17/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------11/17/2021
'--comment Code-----------------------------------------------------------------11/17/2021
'--Update Changelog for release/update------------------------------------------11/17/2021
'--Remove testing message boxes-------------------------------------------------11/17/2021
'--Remove testing code/unnecessary code-----------------------------------------11/17/2021
'--Review/update SharePoint instructions----------------------------------------11/17/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/17/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
