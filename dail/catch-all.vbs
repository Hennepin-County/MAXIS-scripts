'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - CATCH ALL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 195          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
'===============================================================================================END FUNCTIONS LIBRARY BLOCK

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/19/2024", "Improved script functionality and details included in CASE/NOTE.", "Mark Riegel, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("11/01/2019", "BUG FIX - resolved error where script was missing the case notes. Script should now case note every time the script is run to completion.", "Casey Love, Hennepin County")
call changelog_update("09/04/2019", "Reworded the TIKL.", "MiKayla Handley, Hennepin County")
call changelog_update("05/01/2019", "Removed the automated DAIL deletion. Workers must go back and delete manually once the DAIL has been acted on.", "MiKayla Handley, Hennepin County")
call changelog_update("04/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK

'-------------------------------------------------------------------------------------------------------THE SCRIPT
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""

EMReadScreen DAIL_type, 4, 6, 6 'read the DAIL msg'
DAIL_type = trim(DAIL_type)
EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

'Enters “X” on DAIL message to open full message. 
Call write_value_and_transmit("X", 6, 3)

'Read full message, including if message needs to be opened
EMReadScreen full_message_check, 36, 24, 2
If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
    EMReadScreen full_message, 61, 6, 20
    full_message = trim(full_message)
    
    'Remove X from dail message
    EMWriteScreen " ", 6, 3
Else
    ' Script reads the full DAIL message so that it can process, or not process, as needed.
    EMReadScreen full_dail_msg_line_1, 60, 9, 5
    EMReadScreen full_dail_msg_line_2, 60, 10, 5
    EMReadScreen full_dail_msg_line_3, 60, 11, 5
    EMReadScreen full_dail_msg_line_4, 60, 12, 5

    If trim(full_dail_msg_line_2) = "" Then full_dail_msg_line_1 = trim(full_dail_msg_line_1)

    full_message = trim(full_dail_msg_line_1 & full_dail_msg_line_2 & full_dail_msg_line_3 & full_dail_msg_line_4)

    'Transmit back to DAIL message
    transmit

End If

EMWriteScreen "S", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 225, "DAIL_type &   MESSAGE PROCESSED"
  GroupBox 5, 5, 270, 55, "DAIL for case #  &  MAXIS_case_number"
  Text 10, 20, 260, 35, full_message
  Text 10, 65, 45, 15, "Date Doc(s) Received:"
  EditBox 70, 65, 40, 15, docs_rcvd_date
  Text 115, 70, 50, 10, "(if applicable)"
  Text 10, 90, 55, 10, "MEMB Number:"
  EditBox 70, 85, 20, 15, memb_number
  Text 10, 110, 50, 10, "Actions taken:"
  EditBox 70, 105, 205, 15, actions_taken
  Text 10, 130, 50, 10, "Verifs needed:"
  EditBox 70, 125, 205, 15, verifs_needed
  Text 10, 150, 45, 10, "Other notes:"
  EditBox 70, 145, 205, 15, other_notes
  CheckBox 10, 165, 110, 10, "Check here if you want to TIKL", TIKL_checkbox
  CheckBox 10, 180, 90, 10, "ECF has been reviewed ", ECF_reviewed
  Text 5, 210, 60, 10, "Worker signature:"
  EditBox 65, 205, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 205, 45, 15
    CancelButton 230, 205, 45, 15
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
		If trim(docs_rcvd_date) <> "" THEN 
            If isdate(docs_rcvd_date) = False then err_msg = err_msg & vbnewline & "* Please enter a valid date that the forms were received."
        End If
		If (isnumeric(memb_number) = False and len(memb_number) > 2) then err_msg = err_msg & vbcr & "* Please Enter a valid member number."
    	If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

EMReadScreen are_we_in_stat, 14, 20, 11
EMReadScreen are_we_at_dail, 4, 2, 48
If are_we_in_stat = "Function: STAT" Then
    PF3
    EMReadScreen are_we_at_dail, 4, 2, 48
    If are_we_at_dail <> "DAIL" Then
        Call back_to_SELF
        EMWriteScreen "        ", 18, 43
        EMWriteScreen MAXIS_case_number, 18, 43
        Call navigate_to_MAXIS_screen("DAIL", "DAIL")
    End If
ElseIf are_we_at_dail <> "DAIL" Then
    Call back_to_SELF
    EMWriteScreen "        ", 18, 43
    EMWriteScreen MAXIS_case_number, 18, 43
    Call navigate_to_MAXIS_screen("DAIL", "DAIL")
End If

IF TIKL_checkbox = 1 then Call create_TIKL("Review case for requested verifications or actions needed: " & verifs_needed & ".", 10, date, False, TIKL_note_text)

Call start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("-" & DAIL_type & " PROCESSED - " & full_message & "-")
CALL write_variable_in_case_note("---")
IF ECF_reviewed = CHECKED THEN CALL write_variable_in_case_note("* Case file has been reviewed.")
If trim(docs_rcvd_date) <> "" Then CALL write_bullet_and_variable_in_case_note("Date Doc(s) Received", docs_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
IF TIKL_checkbox = CHECKED THEN CALL write_variable_in_case_note(TIKL_date_text)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report(DAIL_type & vbcr & full_message & vbcr & " DAIL has been case noted")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------