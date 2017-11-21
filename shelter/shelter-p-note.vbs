'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "P-NOTE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("11/17/2017", "Updated dialog as requested by Shelter Team", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Custom function not in the FuncLib
Function write_editbox_in_person_note(x, y) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array 
    EMGetCursor row, col 
    If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 18 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(5)
    If (row = 5 and col = 3) then EMSendKey space(5)
    EMSendKey x & " "
    If right(x, 1) = ";" then 
      EMSendKey "<backspace>" & "<backspace>" 
      EMGetCursor row, col 
      If row = 18 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(5)
      Else
        EMSendKey "<newline>" & space(5)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_new_line_in_person_note(x)
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80 + 1 ) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog pnote_dialog, 0, 0, 311, 180, "P-NOTE"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 230, 5, 75, 15, shelter_stay_dates
  EditBox 100, 30, 75, 15, EA_date
  EditBox 100, 50, 75, 15, ACF_date
  EditBox 285, 25, 20, 15, number_nights
  EditBox 285, 45, 20, 15, number_tokens
  EditBox 285, 65, 20, 15, number_buscards
  EditBox 105, 85, 200, 15, reason_for_homelessness
  EditBox 105, 110, 200, 15, resolution_reason
  EditBox 105, 135, 200, 15, other_notes
  Text 10, 10, 45, 10, "Case number:"
  Text 155, 10, 70, 10, "Dates of shelter stay:"
  Text 10, 35, 80, 10, " EA Dates (if applicable):"
  Text 10, 55, 85, 10, " ACF Dates (if applicable):"
  Text 225, 30, 60, 10, "Number of nights:"
  Text 205, 50, 80, 10, "Number of bus token(s):"
  Text 210, 70, 75, 10, "Number of bus card(s):"
  Text 10, 85, 85, 15, "Funds issued when client became homeless due to:"
  Text 60, 115, 40, 10, "Resolution:"
  Text 60, 140, 40, 10, "Other notes:"
  ButtonGroup ButtonPressed
    OkButton 200, 160, 50, 15
    CancelButton 255, 160, 50, 15
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog pnote_dialog
		cancel_confirmation
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF number_nights = "" then err_msg = err_msg & vbNewLine & "* Enter the number of nights of shelter"
		IF shelter_stay_dates = "" then err_msg = err_msg & vbNewLine & "* Please enter the dates of shelter"
		IF reason_for_homelessness = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for homelessness."
		IF resolution = "" then err_msg = err_msg & vbNewLine & "* Enter the resolution."		
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 	
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46
Call navigate_to_MAXIS_screen("STAT", "MEMB")

'Getting the person note ready 
PF5			'navigates to Person note from WREG PANEL
'adds case to the rejected list if cannot person note
EMReadScreen person_note_confirmation, 12, 2, 31
If person_note_confirmation <> "Person Notes" then 
    script_end_procedure ("Person notes cannot be accessed. Please check the case number and servicing county.")
ELSE 
    EMreadscreen edit_mode_required_check, 6, 5, 3		'if not person not exists, person note goes directly into edit mode
    If edit_mode_required_check <> "      " then PF9
        
    'writes the information into the person note
    Call write_new_line_in_person_note("### P-note at End of EA and ACF Shelter Stay ###")
    Call write_editbox_in_person_note("Dates of Shelter Stay", shelter_stay_dates)
    Call write_editbox_in_person_note("Number of Nights in Shelter", number_nights)
    Call write_editbox_in_person_note("Number of bus cards", number_buscards)
    Call write_editbox_in_person_note("Number of Tokens", number_tokens)
    Call write_editbox_in_person_note("ACF Dates (if applicable)", ACF_dates)
    Call write_editbox_in_person_note("EA Dates (if applicable)", EA_dates)
    Call write_editbox_in_person_note("Funds issued when client became homeless due to", reason_for_homelessness)
    Call write_editbox_in_person_note("Resolution", resolution_reason)
    Call write_editbox_in_person_note("Other notes", other_notes)
	Call write_editbox_in_person_note("---")
	Call write_editbox_in_person_note(worker_signature)
	Call write_editbox_in_person_note("Hennepin County Shelter Team")
END IF

script_end_procedure("")