'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

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
BeginDialog pnote_dialog, 0, 0, 316, 105, "P-NOTE"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 235, 5, 75, 15, ACF_EA_dates
  EditBox 60, 25, 20, 15, number_nights
  EditBox 175, 25, 20, 15, number_tokens
  EditBox 175, 45, 135, 15, reason_for_homelessness
  EditBox 55, 65, 255, 15, resolution
  EditBox 55, 85, 145, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 205, 85, 50, 15
    CancelButton 260, 85, 50, 15
  Text 10, 10, 45, 10, "Case number:"
  Text 180, 10, 55, 10, " ACF/EA Dates:"
  Text 15, 90, 40, 10, "Other notes:"
  Text 5, 50, 170, 10, "Funds issued when client become Homeless due to:"
  Text 200, 30, 80, 10, "# bus tokens/bus cards"
  Text 85, 30, 65, 10, "# nights shelter"
  Text 15, 70, 40, 10, "Resolution:"
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
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If number_nights = "" then err_msg = err_msg & vbNewLine & "* Enter the nubmer of nights of shelter"
		If number_tokens = "" then err_msg = err_msg & vbNewLine & "* Enter the number of tokens or buscards"
		If ACF_EA_dates = "" then err_msg = err_msg & vbNewLine & "* Enter the ACF/EA dates."
		If reason_for_homelessness = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for homelessness."
		If resolution = "" then err_msg = err_msg & vbNewLine & "* Enter the resolution."		
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
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
    Call write_editbox_in_person_note("Nights shelter", number_nights)
    Call write_editbox_in_person_note("Tokens or bus cards", number_tokens)
    Call write_editbox_in_person_note("ACF/EA Dates", ACF_EA_dates)
    Call write_editbox_in_person_note("Funds issued when client become Homeless due to", reason_for_homelessness)
    Call write_editbox_in_person_note("Resolution", resolution)
    Call write_editbox_in_person_note("Other notes", other_notes)
	Call write_editbox_in_person_note("---")
	Call write_editbox_in_person_note(worker_signature)
	Call write_editbox_in_person_note("Hennepin County Shelter Team")
END IF

script_end_procedure("")