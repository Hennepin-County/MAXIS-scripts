'Required for statistical purposes===============================================================================
name_of_script = "DAIL - ES REFERRAL MISSING.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
'END OF stats block==============================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.
'DIALOG=====================================================================================================================
BeginDialog ES_ref_dialog, 0, 0, 301, 100, "ES Referral Date"
  EditBox 65, 5, 120, 15, client_name
  EditBox 265, 5, 25, 15, ref_numb
  EditBox 65, 25, 80, 15, ES_ref_date
  CheckBox 5, 45, 205, 10, "Check here to have the script fill this referral date on EMPS", update_emps_checkbox
  EditBox 50, 60, 240, 15, other_notes
  EditBox 70, 80, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 80, 50, 15
    CancelButton 240, 80, 50, 15
  Text 5, 10, 60, 10, "Name from DAIL:"
  Text 195, 10, 70, 10, "HH member number:"
  Text 5, 30, 60, 10, "ES Referral Date:"
  Text 155, 30, 140, 10, "If filled, date was gathered on INFC/WORK"
  Text 5, 65, 40, 10, "Other notes:"
  Text 5, 85, 60, 10, "Worker signature:"
EndDialog
'===========================================================================================================================

EMConnect ""		'Getting Case number
EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

EMReadScreen name_for_dail, 57, 5, 5			'Reading the name of the client
'This next block will determine the name of the client the message is for
'If the message is for someone other than M01 - the name is writen next to the name of M01
other_person = InStr(name_for_dail, "--(")	'This determines if it for someone other than M01
'This is for if the message is for M01'
If other_person = 0 Then
	comma_loc = InStr(name_for_dail, ",")  	'Determines the end of the last name
	dash_loc = InStr(name_for_dail, "-")	'Determines the end of the name
	EMReadscreen last_name, comma_loc - 1, 5, 5									'Reading the last name
	EMReadscreen middle_exists, 1, 5, 5 + (dash_loc - 2)						'Determines if clt's middle initial is listed
	If middle_exists = " " Then 												'If not - reads first name
		EMReadscreen first_name, dash_loc - comma_loc - 5, 5, comma_loc + 5
	Else 																		'If so - reads first name
		EMReadScreen first_name, dash_loc - comma_loc - 3, 5, comma_loc + 5
	End If
'This is for if the message is for a different HH Member
Else
	end_other = InStr(name_for_dail, ")--")
	comma_loc = InStr(other_person, name_for_dail, ",")
	EMReadscreen last_name, comma_loc - other_person - 3, 5, other_person + 7
	EMReadscreen middle_exists, 1, 5, end_other + 2
	If middle_exists = " " Then
		EMReadscreen first_name, end_other - comma_loc - 3, 5, comma_loc + 5
	Else
		EMReadScreen first_name, end_other - comma_loc - 1, 5, comma_loc + 5
	End If
End If
client_name = last_name & ", " & first_name		'putting the name into one string

'Going to INFC WORK
EMSendKey "i"
transmit

EMSendKey "work"
transmit

'Confirming we are at WORK
EMReadScreen work_panel_check, 4, 2, 51
If work_panel_check = "WORK" Then
work_maxis_row = 7
	DO
		EMReadScreen work_name, 26, work_maxis_row, 7			'Reads the client name from INFC/WORK'
		work_name = trim(work_name)
		IF client_name = work_name then
			memb_check = vbYes		'If the name on INFC/WORK exactly matches the name from the DAIL, the script does not need user input and will gather the Reference Number'
			EMReadScreen ref_numb, 2, work_maxis_row, 3
		ElseIf client_name <> work_name then 	'if name doesn't match the referral name the confirmation is required by the user
			memb_check = MsgBox ("DAIL Message is for - " & client_name & vbNewLine & "Name on INFC/WORK - " & work_name & _
			  vbNewLine & vbNewLine & "Is this the client you need ES Referral Information about?", vbYesNo + vbQuestion, "Confirm Client using Banked Monhts")
			If memb_check = vbYes Then		'If the user confirms that this is the correct client, Ref number is gathered'
				EMReadScreen ref_numb, 2, work_maxis_row, 3
			ElseIf memb_check = vbNo Then	'If the user says NO the script will see if there are other clients listed on INFC/WORK and start back at the beginning of the loop to try to match'
				EMReadScreen next_clt, 1, (work_maxis_row + 1), 7
			END IF
		End If
		work_maxis_row = work_maxis_row + 1		'Increments to read the next row for a new client'
		STATS_counter = STATS_counter + 1
	Loop until next_clt = " " OR memb_check = vbYes

	'Reads the referral date from WORK for the client found above
	If memb_check = vbYes Then EMReadScreen es_ref_date, 8, 7, 72
	If es_ref_date = "__ __ __" Then es_ref_date = ""
	es_ref_date = replace(es_ref_date, " ", "/")
End If

PF3 	'Back to DAIL
EMWriteScreen "s", 6, 3			'Goes to EMPS
transmit
EMWriteScreen "emps", 20, 71
transmit

If ref_numb <> "" Then 		'gets to the right EMPS panel
	EMWriteScreen ref_numb, 20, 76
	transmit
End If

'Defaulting to having the script update EMPS
update_emps_checkbox = checked

'Runs the dialog
Do
	err_msg = ""
	Dialog ES_ref_dialog
	cancel_confirmation
	If worker_signature = "" Then err_msg = err_msg & vbNewLine & "Sign your case note."
	If isdate(es_ref_date) = FALSE Then err_msg = err_msg & vbNewLine & "You must enter a valid date for the ES Referral Date."
	If update_emps_checkbox = checked AND es_ref_date = "" Then err_msg = err_msg & vbNewLine & "You must have a date entered for the script to update EMPS"
	If update_emps_checkbox = checked AND ref_numb = "" Then err_msg = err_msg & vbNewLine & "You must enter the client's reference number in order for the EMPS panel to be correctly updated."
	If err_msg <> "" Then MsgBox "Please resolve before you continue." & vbNewLine & err_msg
Loop until err_msg = ""

'If the user requests the script to update
If update_emps_checkbox = checked Then
	Call Navigate_to_MAXIS_screen ("STAT", "EMPS")		'Makes sure we are still at EMPS
	EMWriteScreen ref_numb, 20, 76						'And at the correct client
	transmit
	PF9													'Edit
	ref_month = right("00" & DatePart("m", es_ref_date), 2)
	ref_date  = right("00" & DatePart("d", es_ref_date), 2)
	ref_year  = right(DatePart("yyyy", es_ref_date), 2)
	EMWriteScreen ref_month, 16, 40						'Write in the date
	EMWriteScreen ref_date,  16, 43
	EMWriteScreen ref_year,  16, 46
	transmit
	PF3

 	'Check to make sure we are back to our dail
 	EMReadScreen DAIL_check, 4, 2, 48
 	IF DAIL_check <> "DAIL" THEN
 		PF3 'This should bring us back from UNEA or other screens
 		EMReadScreen DAIL_check, 4, 2, 48
 		IF DAIL_check <> "DAIL" THEN 'If we are still not at the dail, try to get there using custom function, this should result in being on the correct dail (but not 100%)
 			call navigate_to_MAXIS_screen("DAIL", "DAIL")
 		END IF
 	END IF
 	EMWriteScreen "n", 6, 3
 	transmit

 	'Case noting
 	PF9
 	EMReadScreen case_note_mode_check, 7, 20, 3
 	If case_note_mode_check <> "Mode: A" then MsgBox "You are not in a case note on edit mode. You might be in inquiry. Try the script again in production."
 	If case_note_mode_check <> "Mode: A" then stopscript

	Call Write_Variable_in_CASE_NOTE ("DAIL Processed - ES Referal Date Updated for Memb " & ref_numb)
	Call Write_Variable_in_CASE_NOTE ("* PEPR message rec'vd indicating that EMPS panel was missing ES Referral Date")
	If memb_check = vbYes Then Call Write_Variable_in_CASE_NOTE ("* ES Referral Date found on INFC/WORK and added to EMPS")
	Call Write_Bullet_and_Variable_in_Case_Note ("Date Entered", es_ref_date)
	Call Write_Bullet_and_Variable_in_Case_Note ("Notes", other_notes)
	Call Write_Variable_in_CASE_NOTE ("---")
	Call Write_Variable_in_CASE_NOTE (worker_signature)
	end_msg = "Success! EMPS has been updated and Case Note Written"
Else
	end_msg = "You have selected to not have the EMPS panel updated by the script." & vbNewLine & "You will need to process this DAIL manually."

End If

script_end_procedure(end_msg)
