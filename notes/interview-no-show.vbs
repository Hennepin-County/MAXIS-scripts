'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INTERVIEW NO SHOW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("03/21/2019", "Updated script to align with the On Demand process. Now for walk-ins only. Removed NOMI options.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
BeginDialog case_number_dialog, 0, 0, 141, 50, "Enter the Case Number"
  EditBox 65, 10, 70, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 65, 30, 35, 15
    CancelButton 100, 30, 35, 15
  Text 10, 15, 50, 10, "Case Number:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""
EMFocus

Call check_for_MAXIS(true)

'Pulls case number from MAXIS if worker has already selected a case
Call MAXIS_case_number_finder(MAXIS_case_number)

Do
    err_msg = ""

    Dialog case_number_dialog

    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "* Enter a case number."
    If len(MAXIS_case_number) > 7 Then err_msg = err_msg & vbNewLine & "* The case number is too long, review"
    If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid MAXIS case number."

    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

Loop Until err_msg = ""

BeginDialog same_day_dialog, 0, 0, 191, 290, "Enter No Show Information"
  Text 70, 25, 60, 15, MAXIS_case_number
  EditBox 70, 75, 90, 15, interview_date
  EditBox 70, 95, 90, 15, first_page
  EditBox 70, 115, 90, 15, second_page
  CheckBox 20, 150, 155, 20, "Attempted phone call to client - No Answer", pc_attempted
  EditBox 75, 170, 95, 15, time_called
  EditBox 75, 190, 95, 15, phone_number
  CheckBox 75, 210, 90, 15, "Left Message for Client", left_vm
  CheckBox 10, 235, 70, 15, "Potential XFS", potential_xfs
  EditBox 70, 250, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 105, 270, 40, 15
    CancelButton 145, 270, 40, 15
  Text 10, 5, 170, 10, "Client did not respond to page for in-office interview"
  Text 15, 25, 45, 10, "Case Number"
  GroupBox 5, 60, 175, 75, "Client was Paged in the Lobby"
  Text 15, 80, 50, 10, "Interview Date:"
  Text 15, 100, 50, 10, "1st Page time:"
  Text 15, 120, 50, 10, "2nd Page time:"
  GroupBox 5, 140, 175, 90, "Phone Call to Client"
  Text 35, 175, 35, 10, "Called at:"
  Text 15, 195, 50, 15, "Phone Number"
  Text 10, 255, 60, 10, "Worker Signature"
  Text 5, 45, 60, 10, "Application Date:"
  EditBox 70, 40, 90, 15, application_date
EndDialog

'Defaults the Interview Date to today's date
interview_date = date & ""

'Defaults the Client Phone number to the first phone number listed on MAXIS in STAT/ADDR
Call navigate_to_MAXIS_screen ("STAT", "ADDR")
EMReadScreen phone_01, 3, 17, 45
EMReadScreen phone_02, 3, 17, 51
EMReadScreen phone_03, 4, 17, 55
phone_number = phone_01 & "-" & phone_02 & "-" & phone_03 & ""


'Determines which programs are currently pending in the month of application
call navigate_to_MAXIS_screen("STAT","PROG")
EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen emer_pend, 4, 8, 74
EMReadScreen grh_pend, 4, 9, 74
EMReadScreen fs_pend, 4, 10, 74
EMReadScreen hc_pend, 4, 12, 74

'Assigns a value so the programs pending will show up in check boxes
IF cash1_pend = "PEND" THEN
	cash1_pend = 1
    EMReadScreen app_date, 8, 6, 33
Else
	cash1_pend = 0
End If

If cash2_pend = "PEND" THEN
	cash2_pend = 1
    EMReadScreen app_date, 8, 7, 33
Else
	cash2_pend = 0
End if

If cash1_pend = 1 OR cash2_pend = 1 then cash_pend = 1

If emer_pend = "PEND" THEN
	emer_pend = 1
    EMReadScreen app_date, 8, 8, 33
Else
	emer_pend = 0
End if

If grh_pend = "PEND" THEN
	grh_pend = 1
    EMReadScreen app_date, 8, 9, 33
Else
	grh_pend = 0
End if

If hc_pend = "PEND" THEN
	hc_pend = 1
    EMReadScreen app_date, 8, 12, 33
Else
	hc_pend = 0
End if

If fs_pend = "PEND" THEN
	fs_pend = 1
    EMReadScreen app_date, 8, 10, 33
Else
	fs_pend = 0
End if

If app_date <> "" AND app_date <> "__ __ __" Then application_date = replace(app_date, " ", "/")

'Display's the Dialog Box to imput variable information - includes safeguards for mandatory fields
Do
	Do
		Do
			err_msg = ""
			Dialog same_day_dialog
			cancel_confirmation
			IF MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
			IF interview_date = "" THEN err_msg = err_msg & vbNewLine & "*Please enter an Interview Date"
			IF IsDate (interview_date) = False THEN err_msg = err_msg & vbNewLine & "*Please enter a valid Interview Date"
			IF first_page = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the time of the 1st page in the lobby"
			IF second_page = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the time of the second page in the lobby - you must page your client at least twice"
			IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*Please sign your case note"
			If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
		Loop until err_msg = ""
		'The following converts the times entered by the user to a standard format
		IF IsNumeric(first_page) = TRUE THEN
			first_page = FormatNumber (first_page, 2)
			first_page = FormatDateTime (first_page, 4)
		End If
		IF IsNumeric(second_page) = TRUE THEN
			second_page = FormatNumber (second_page, 2)
			second_page = FormatDateTime (second_page ,4)
		End If
		first_page = TimeValue(first_page)
		second_page = TimeValue(second_page)
		'This converts the time to military time for any afternnon times
		If first_page < TimeValue("7:00") THEN first_page = DateAdd("h", 12, first_page)
		If second_page < TimeValue("7:00") THEN second_page = DateAdd("h", 12, second_page)
		'This tests to ensure the page times are at least 15 minutes apart
		IF abs(DateDiff("n", first_page, second_page))<15 THEN MsgBox "You must page client at least 15 minutes apart"
	Loop until abs(DateDiff("n", first_page, second_page))>=15 'and MAXIS_case_number <> "" and interview_date <> "" and IsDate(interview_date) = True and first_page <> "" and second_page <> "" and worker_signature <> ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

call check_for_MAXIS(False)

'Formats the page times and time called to standard hh:mm for case note
first_page = FormatDateTime (first_page, 4)
second_page = FormatDateTime (second_page ,4)
IF IsNumeric(time_called) = TRUE THEN
	time_called = FormatNumber (time_called, 2)
	time_called = FormatDateTime (time_called, 4)
End If

'Creates a variable that lists all the programs pending.
If cash_pend = 1 THEN programs_applied_for = programs_applied_for & "Cash, "
If emer_pend = 1 THEN programs_applied_for = programs_applied_for & "Emergency, "
If grh_pend = 1 THEN programs_applied_for = programs_applied_for & "GRH, "
If fs_pend = 1 THEN programs_applied_for = programs_applied_for & "SNAP, "
If hc_pend = 1 THEN programs_applied_for = programs_applied_for & "HC, "

If programs_applied_for = "" Then programs_applied_for = "None pending in MAXIS at this time"
programs_applied_for = left(programs_applied_for, len(programs_applied_for) -2)

'Starts a Case Note
Call start_a_blank_case_note

call write_variable_in_CASE_NOTE("*** Attempted to Page Client in Lobby for Interview - No Show ***")
call write_bullet_and_variable_in_CASE_NOTE("Date of application", application_date)
call write_bullet_and_variable_in_CASE_NOTE("Client walked in to office for interview", interview_date)
call write_bullet_and_variable_in_CASE_NOTE("Paged client in the lobby to complete interview at", first_page & " & " & second_page)
IF pc_attempted = 1 THEN call write_bullet_and_variable_in_CASE_NOTE("Attempted to call client, no answer, called at provided number", phone_number & " at " & time_called)
IF left_vm = 1 THEN call write_variable_in_CASE_NOTE("* Left Voicemail for Client.")
IF potential_xfs = 1 THEN call write_variable_in_CASE_NOTE("* Case is Potentially XFS")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure ("")
