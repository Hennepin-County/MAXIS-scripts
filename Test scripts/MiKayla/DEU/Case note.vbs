'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CASE NOTE TO EMAIL"
start_time = timer
STATS_counter = 0              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.
STATS_manualtime = 160          'manual run time in seconds
STATS_denomination = "C"       'I is for each dail message

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
'----------------------------------------------------------------------------------------------Dialog
BeginDialog case_note, 0, 0, 251, 95, "Case note"
  EditBox 65, 10, 75, 15, MAXIS_case_number
  EditBox 65, 30, 75, 15, casenote_date
  EditBox 65, 50, 180, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 135, 75, 50, 15
    CancelButton 195, 75, 50, 15
  Text 10, 55, 45, 10, "Other Notes:"
  Text 10, 35, 55, 10, "Case Note Date:"
  Text 10, 15, 50, 10, "Case Number: "
EndDialog


'Connects to MAXIS
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Do
	dialog case_note
    cancel_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


Call navigate_to_MAXIS_screen("CASE", "NOTE")

EMWriteScreen "x", 5, 3
Transmit
note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 76, note_row, 3
        note_line = trim(note_line)
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & vbcr		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
			EMReadScreen next_page, 7, note_row, 3
			If next_page = "More: +" Then 						'This indicates there is another page of the case note
				PF8												'goes to the next line and resets the row to read'\
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages

    msgbox message_array
    'MsgBox note_header & vbcr & note_line_two & vbcr & note_line_three & vbcr & note_line_four & vbcr & note_line_five & vbcr & note_line_six & vbcr & note_line_seven & vbcr & note_line_eight & vbcr & note_line_nine & vbcr & note_line_ten & vbcr & note_line_eleven & vbcr & note_line_twelve & vbcr & note_line_thirteen & vbcr & note_line_fourteen
    'IF programs = "Health Care" or programs = "Medical Assistance" THEN
    'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
    CALL create_outlook_email("mikayla.handley@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & OP_Date & "Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)

script_end_procedure("Success!")
