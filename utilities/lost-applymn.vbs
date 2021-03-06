name_of_script = "UTILITIES - Lost ApplyMN.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
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
call changelog_update("03/10/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
function search_for_info_in_note(date_of_app, applymn_confirmation_number, name_of_applicant)
	'This funciton is to specifically search case notes for an application received case note that gives us date of application, confirmation number and client name
	four_months_ago = DateAdd("m", -4, date)		'we really shouldn't have things older than 4 months here
	If MAXIS_case_number <> "" Then					'The Case Number is needed to navigate
		Call navigate_to_MAXIS_screen("CASE", "NOTE")							'First fgo to Case Note
		EMReadScreen name_of_applicant, 26, 21, 40								'Read the applicant's name from be bottom of the case note screen
		name_of_applicant = trim(name_of_applicant)								'trim extra spaces
		name_of_applicant = replace(name_of_applicant, ",", ", ")				'Put a space after the comma for readability

		note_row = 5															'Case notes start at row 5
		Do
			EMReadScreen case_note_header, 32, note_row, 25						'read the header and the date of the note
			EMReadScreen case_note_date, 8, note_row, 6
			' MsgBox "~" & case_note_date & "~"

			case_note_header = trim(case_note_header)							'reformat the case note
			If case_note_date <> "        " Then case_note_date = DateAdd("d", 0, case_note_date)		'make the date a date

			If case_note_header ="~ Application Received (ApplyMN)" Then		'If it finds an application received case note for an ApplyMN
				EmWriteScreen "X", note_row, 3									'open the case note
				transmit

				EMReadScreen date_of_app, 10, 4, 50								'read the app date and reformat
				date_of_app = replace(date_of_app, "~", "")
				date_of_app = trim(date_of_app)

				the_row = 1														'search for the confirmation field
				the_col = 1
				EMSearch "* Confirmation #", the_row, the_col
				EMReadScreen applymn_confirmation_number, 30, the_row, 22		'Read the confirmation number
				applymn_confirmation_number = trim(applymn_confirmation_number)	'reformat
				PF3																'leave the case note
				Exit Do															'we don't need to search any other case notes
			End If

			note_row = note_row + 1					'go to the next note
			If note_row = 19 Then					'if we are at the bottom of the page, go to the next page and start at the top
				PF8
				note_row = 5
				EMReadScreen end_of_list, 9, 24, 14
				If end_of_list = "LAST PAGE" Then exit Do
			End If
		Loop until case_note_header = "" OR case_note_date < four_months_ago	'stop looking if no more notes or the notes are more than 4 months old.
	End If
end function
'===========================================================================================================================
'Connecting to BlueZone
EMConnect ""

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
Call check_for_MAXIS(True)								'make sure we are in MAXIS
CALL MAXIS_case_number_finder (MAXIS_case_number)		'try to find the case number

'One and only dialog for this script
DO
    DO
        err_msg = ""

		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 331, 165, "ApplyMN not Found"
		  EditBox 140, 35, 80, 15, MAXIS_case_number
		  CheckBox 25, 55, 225, 10, "Check here if this is for a NEW request with no Case Number yet.", no_case_number_checkbox
		  EditBox 80, 100, 50, 15, date_of_app
		  EditBox 225, 100, 75, 15, applymn_confirmation_number
		  EditBox 80, 120, 220, 15, name_of_applicant
		  EditBox 65, 145, 115, 15, worker_name
		  ButtonGroup ButtonPressed
		    PushButton 165, 80, 135, 15, "Read ApplyMN Info from CASE:NOTE", collect_from_case_note_btn
		    OkButton 215, 145, 50, 15
		    CancelButton 270, 145, 50, 15
		  Text 10, 10, 315, 20, "If a client is reporting they have submitted an ApplyMN application, and there is no coresponding application in ECF, this script can assist in sending the request to QI to find the ApplyMN."
		  Text 15, 40, 120, 10, "Case Number with the lost ApplyMN:"
		  GroupBox 10, 70, 310, 70, "ApplyMN Detail"
		  Text 15, 105, 65, 10, "Date of Application:"
		  Text 150, 105, 75, 10, "Confirmation Number:"
		  Text 25, 125, 55, 10, "Applicant Name: "
		  Text 10, 150, 55, 10, "Sign your Email"
		  ButtonGroup ButtonPressed
		    PushButton 300, 80, 15, 15, "?", read_note_question
		    PushButton 300, 100, 15, 15, "?", confirmation_number_question
		    PushButton 300, 120, 15, 15, "?", applicant_name_question
		    PushButton 130, 100, 15, 15, "?", app_date_question
		EndDialog

        Dialog Dialog1
        cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")
		If no_case_number_checkbox = checked then err_msg = ""			'if the checkbox is check it will blank out the case number error messaging

		If IsDate(date_of_app) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of application as a valid date."
		applymn_confirmation_number = trim(applymn_confirmation_number)
		If applymn_confirmation_number = "" Then err_msg = err_msg & vbNewLine & "* Enter the confirmation number for the ApplyMN."
		name_of_applicant = trim(name_of_applicant)
		If name_of_applicant = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the person who submitted the ApplyMN."
		worker_name = trim(worker_name)
		If worker_name = "" Then err_msg = err_msg & vbNewLine & "* Enter your name to sign your email."

		If ButtonPressed = collect_from_case_note_btn Then
			err_msg = "LOOP" & err_msg
			Call search_for_info_in_note(date_of_app, applymn_confirmation_number, name_of_applicant)		'Ths will call the in script function to fill in some of the fields
		End If

		If ButtonPressed = read_note_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("When can you use the functionality to read ApplyMN Information from CASE:NOTE and what will it read?"& vbNewLine & vbNewLine & "*** THE NOTE THIS REFERS TO IS THE APPLICATION RECEIVED NOTE WITH APPLYMN INFORMATION. ****" & vbNewLine & vbNewLine & " - If the case was pended and the script NOTES - Application Received is used to capture the information then the details are listed in CASE:NOTE." & vbNewLine & " - What can be read:" & vbNewLine & "    *Application Date" & vbNewLine & "    *Confirmation Number" & vbNewLine & "    *Name of Applicant", vbInformation, "What can e read from CASE:OTE")
		End If
		If ButtonPressed = confirmation_number_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("The confirmation number is the primary means to search through the files of ApplyMNs submitted" & vbNewLine & vbNewLine & "There is no other information that allows us to search as quickly for ApplyMNs." & vbNewLine & vbNewLine & "--- What if the client does not know their Confirmation Number?" & vbNewLine & "Information from when the client submitted the application online can be found from logging back into their account for ApplyMN.", vbInformation, "Why do we need the confirmation number?")
		End If
		If ButtonPressed = applicant_name_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("The application name allows us to most easily confirm the application is correct as it is unique identifying information. Enter this information as it was entered in the application if possible.", vbInformation, "Why is the applicant name needed?")
		End If
		If ButtonPressed = app_date_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("There are hundreds of ApplyMN applications submitted in any day. This number has increased more recently." & vbNewLine & vbNewLine & "*** The date needs to be the precise date the application was submitted ***" & vbNewLine & vbNewLine & "This is important because the appplications must be pulled by date before searching. Trying to pull either the wrong date or mutliple dates prevents us from finding this information." & vbNewLine & vbNewLine & "Information from when the client submitted the application online can be found from logging back into their account for ApplyMN.", vbInformation, "Why is the application date needed?")
		End If
		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "Please resolve to continue: " & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Now we go look to see if the case is pending
case_appld = FALSE
If no_case_number_checkbox = unchecked Then
	Call navigate_to_MAXIS_screen("CASE", "CURR")		'Go to CASE CURR

	curr_row = 1
	curr_col = 2
	EMSearch "PENDING", curr_row, Curr_col				'find the word 'PENDING" anywhere and assume it has been APPLd and is pending

	If curr_row <> 0 Then case_appld = TRUE
End If

email_subject = "Request for Search for ApplyMN not in ECF"			'Setting the subject line

email_body = "Please recover the ApplyMN file that cannot be found." & vbCr & vbCr						'Fillin i nthe information for the body of the email

If no_case_number_checkbox = unchecked Then email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr
If no_case_number_checkbox = checked Then email_body = email_body & "No Case Number known at this time as case has not been pended." & vbCr
email_body = email_body & "Name of Applicant: " & name_of_applicant & vbCr
email_body = email_body & "Date of Application: " & date_of_app & vbCr
email_body = email_body & "Confirmation Number: " & applymn_confirmation_number & vbCr

If case_appld = TRUE Then email_body = email_body & vbCr & "~~The case has been APPL'd. ~~" & vbCr
If case_appld = FALSE Then email_body = email_body & vbCr & "~~This case is not pending and has not been APPL'd.~~" & vbCr

email_body = email_body & vbCr & "Thank you, " & vbCr & worker_name

Call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", TRUE)		'Send the Email

'Add a message with the email information for display
end_msg = "Success!" & vbNewLine & vbNewLine
end_msg = end_msg & "Your email has been sent to QI. Email sent:" & vbCr
end_msg = end_msg & "--------------------------------------------------" & vbCr & vbCr
end_msg = end_msg & email_body

call script_end_procedure_with_error_report(end_msg)
