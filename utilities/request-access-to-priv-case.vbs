name_of_script = "UTILITIES - Request Access to PRIV Case.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 45                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("08/19/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'===========================================================================================================================
'Connecting to BlueZone
EMConnect ""

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
' Call check_for_MAXIS(True)								'make sure we are in MAXIS
CALL MAXIS_case_number_finder (MAXIS_case_number)		'try to find the case number
EMReadScreen SELF_check, 4, 2, 50		'Does this to check to see if we're on SELF screen
IF SELF_check = "SELF" THEN				'if on the self screen then x # is read from coordinates
	EMReadScreen x_number, 7, 22, 8
End If
If x_number = "" Then x_number = "x127"

'One and only dialog for this script
DO
	email_body = ""
	email_subject = ""
    DO
		err_msg = ""

		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 306, 110, "PRIV Case Access"
		  EditBox 80, 25, 60, 15, MAXIS_case_number
		  EditBox 80, 45, 60, 15, x_number
		  EditBox 80, 65, 200, 15, notes
		  EditBox 80, 90, 115, 15, worker_name
		  ButtonGroup ButtonPressed
		    OkButton 200, 90, 50, 15
		    CancelButton 255, 90, 50, 15
		  Text 10, 10, 280, 10, "Request Knowledge Now to update MAXIS to allow you access to a privileged case."
		  Text 10, 30, 70, 10, "PRIV Case Number:"
		  Text 20, 50, 55, 10, "Your X-Number:"
		  Text 15, 70, 60, 10, "Information/Notes:"
		  Text 20, 95, 55, 10, "Sign your Email"
		EndDialog


        Dialog Dialog1
        cancel_without_confirmation

		MAXIS_case_number = trim(MAXIS_case_number)
		x_number = trim(x_number)

		Call validate_MAXIS_case_number(err_msg, "*")
		If len(x_number) <> 7 Then err_msg = err_msg & vbNewLine & "* Review the worker number entered, it is not the right length"
		If ucase(left(x_number, 4)) <> "X127" Then err_msg = err_msg & vbNewLine & "* Review the worker number entered, it does not start with 'x127'."

		If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = ""

	email_subject = "PRIV Case Access Request"

	notes = trim(notes)
	worker_name = trim(worker_name)

	email_body = "Please update MAXIS to allow access to this privileged case." & vbCr & vbCr

	email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr
	email_body = email_body & "Worker Number for transfer: " & x_number & vbCr & vbCr

	If notes <> "" Then email_body = email_body & "Notes: " & notes & vbCr & vbCr
	email_body = email_body & "---" & vbCr
	If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

	message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL TO KNOWLEDGE NOW:" & vbCr & vbCr & email_subject & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)
Loop until message_confirmed = vbYes


email_body = "~~This email is generated from completion of the 'Request Access to PRIV Case' Script.~~" & vbCr & vbCr & email_body
call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", TRUE)

STATS_manualtime = STATS_manualtime + (timer - start_time)
end_msg = "Thank you!" & vbCr & "Your request for access has been sent to QI Knowledge Now." & vbCr & vbCr
end_msg = end_msg & "Content of your Email to Knowledge Now:" & vbCr & "----------------------------------------------------------" & vbCr
end_msg = end_msg & "Subject: " & email_subject & vbCr & vbCr
end_msg = end_msg & email_body

call script_end_procedure_with_error_report(end_msg)
