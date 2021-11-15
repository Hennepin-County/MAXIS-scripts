name_of_script = "UTILITIES - Request Access to PRIV Case.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 90                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
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
call changelog_update("08/19/2020", "PRIV Case Access script will now review for Foster Care and Safe at Home restricted baskets to provide the correct email action for the case entered. Additionally, script will no longer send an email if it is indicated that the resident is on the phone, these requests are more timely when completed in Teams.", "Casey Love, Hennepin County")
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
		  CheckBox 160, 30, 125, 10, "Check here if you are on the phone", resident_on_phone_checkbox
		  EditBox 80, 45, 60, 15, x_number
		  EditBox 80, 65, 200, 15, notes
		  EditBox 80, 90, 115, 15, worker_name
		  ButtonGroup ButtonPressed
		    OkButton 200, 90, 50, 15
		    CancelButton 255, 90, 50, 15
		  Text 10, 10, 280, 10, "Request Knowledge Now to update MAXIS to allow you access to a privileged case."
		  Text 10, 30, 70, 10, "PRIV Case Number:"
		  Text 170, 40, 60, 10, "with the resident."
		  Text 20, 50, 55, 10, "Your X-Number:"
		  Text 15, 70, 60, 10, "Information/Notes:"
		  Text 20, 95, 55, 10, "Sign your Email"
		EndDialog

        Dialog Dialog1					'displaying the dialog
        cancel_without_confirmation

		x_number = trim(x_number)

		Call validate_MAXIS_case_number(err_msg, "*")
		If len(x_number) <> 7 Then err_msg = err_msg & vbNewLine & "* Review the worker number entered, it is not the right length"
		If ucase(left(x_number, 4)) <> "X127" Then err_msg = err_msg & vbNewLine & "* Review the worker number entered, it does not start with 'x127'."

		If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = ""

	Call back_to_SELF								'trying to get in to the case in STAT
	Call navigate_to_MAXIS_screen("STAT", "SUMM")

	EMReadScreen priv_worker_x_number, 7, 24, 65	'reading the x number on the message if access isn't allowed.

	'These are Safe at Home cases
	If priv_worker_x_number = "X127966" OR priv_worker_x_number = "X127AP7" OR priv_worker_x_number = "X127Q95" OR priv_worker_x_number = "X127FAT" Then
		priv_case_type = "Safe at Home"
		If priv_worker_x_number = "X127966" Then
			priv_case_worker_name = "Florence Manley"
			priv_case_worker_email = "Florence.Manley@hennepin.us"
		End If
		If priv_worker_x_number = "X127AP7" Then
			priv_case_worker_name = "Ryan Kierth"
			priv_case_worker_email = "Ryan.Kierth@hennepin.us"
		End If
		If priv_worker_x_number = "X127Q95" Then
			priv_case_worker_name = "Shanna Hansen"
			priv_case_worker_email = "Shanna.Hansen@hennepin.us"
		End If
		If priv_worker_x_number = "X127FAT" Then
			priv_case_worker_name = "See Xiong"
			priv_case_worker_email = "See.Xiong@hennepin.us"
		End If
	End If
	'These are foster care cases
	If priv_worker_x_number = "x127FG1" OR priv_worker_x_number = "x127FG2" OR priv_worker_x_number = "x127EW4" Then
		priv_case_type = "Foster Care"
		priv_case_worker_name = "Team 469"
		priv_case_worker_email = "hsph.es.team.469@hennepin.us"
	End If

	'If a privileged case type was identified there is special handling to send an email to the correct worker/team
	If priv_case_type <> "" Then
		STATS_manualtime = STATS_manualtime + 120 								'adding time for review of restricted baskets and process as the script has completed this.
		Do
			BeginDialog Dialog1, 0, 0, 306, 165, "Case is in a restricted basket"
			  EditBox 10, 70, 290, 15, case_contact_reason
			  EditBox 10, 100, 290, 15, notes
			  ButtonGroup ButtonPressed
			    PushButton 170, 125, 130, 15, "Send an Email about this Case", send_email_to_team_btn
			    PushButton 170, 145, 130, 15, "No Email Needed", no_email_button
			  Text 10, 10, 295, 10, "This case is privileged and is not transferred outside of the team that works on the cases."
			  Text 25, 25, 60, 10, "PRIV Case type:"
			  Text 85, 25, 85, 10, priv_case_type
			  Text 20, 40, 65, 10, "PRIV Case worker:"
			  Text 85, 40, 85, 10, priv_case_worker_name
			  Text 10, 60, 70, 10, "Reason for Contact:"
			  Text 10, 90, 95, 10, "Additional Notes for Email:"
			EndDialog

			dialog Dialog1			'showing the dialog

			'This option will send an email to the worker/team about the case.
			If ButtonPressed = send_email_to_team_btn Then
				email_subject = "PRIV Case Contact"

				notes = trim(notes)
				worker_name = trim(worker_name)
				case_contact_reason = trim(case_contact_reason)

				email_body = "Hello! " & priv_case_worker_name & vbCr & "Privileged MAXIS case contact." & vbCr & vbCr

				email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr & vbCr
				If case_contact_reason <> "" Then email_body = email_body & "Contact Reason: " & case_contact_reason & vbCr & vbCr
				If notes <> "" Then email_body = email_body & "Notes: " & notes & vbCr & vbCr
				email_body = email_body & "---" & vbCr
				If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

				message_confirmed = MsgBox("REVIEW THE WORDING OF YOUR EMAIL ABOUT THIS PRIVILEGED CASE:" & vbCr & vbCr & email_subject & vbCr & vbCr & email_body, vbQuestion + vbYesNo, email_subject)

				end_msg = "Thank you!" & vbCr & "Your request for access has been sent to " & priv_case_worker_name & "." & vbCr & vbCr
				end_msg = end_msg & "Content of your Email:" & vbCr & "----------------------------------------------------------" & vbCr
				end_msg = end_msg & "Subject: " & email_subject & vbCr & vbCr
				end_msg = end_msg & email_body
			End If

			If ButtonPressed = no_email_button Then
				end_msg = "Script run is complete. Case is Privileged because it is a " & priv_case_type & " case." & vbCr & vbCr
				end_msg = end_msg & "Case is privileged to " & priv_case_worker_name & "." & vbCr & vbCr
				end_msg = end_msg & "No email sent regarding this case."
			End If
		Loop until message_confirmed = vbYes OR ButtonPressed = no_email_button
		If ButtonPressed = send_email_to_team_btn Then
			email_body = "~~This email is generated from completion of the 'Request Access to PRIV Case' Script.~~" & vbCr & vbCr & email_body
			call create_outlook_email(priv_case_worker_email, "", email_subject, email_body, "", TRUE)
			STATS_manualtime = STATS_manualtime + (timer - start_time)						'This script allows for the writing of the email - so the manual time is adjusted as email length will vary
		End If
		call script_end_procedure_with_error_report(end_msg)			'End the script run here because if the case is in one of these types, we do not need to request from Knowledge Now
	End If

	'If the user checked the box that they are on the phone, the script brings up this dialog to recommend to use Teams for the request.
	If resident_on_phone_checkbox = checked Then
		BeginDialog Dialog1, 0, 0, 191, 110, "Request Access to Case via Teams"
		  ButtonGroup ButtonPressed
		    PushButton 55, 70, 130, 15, "View Knowledge Now Page", open_knowledge_now_btn
		    PushButton 55, 90, 130, 15, "Complete Script Run", somplete_script_btn
		  Text 10, 10, 175, 20, "Requests for access to a PRIV Case while on the phone with the resident is best completed via Teams"
		  Text 10, 40, 155, 10, "This script can only request access via Email."
		  Text 15, 50, 175, 10, "Email requests are ideal for other case processing."
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_knowledge_now_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/Lists/Knowledge%20Now/calendar.aspx"

		end_msg = "Script Run complete." & vbCr & vbCr & "No Email sent." & vbCr & vbCr & "Please reach out via Teams to Knowledge Now for access to the case."
		call script_end_procedure_with_error_report(end_msg)			'ending the script run here since we should not continue to the email.
	End If

	email_subject = "PRIV Case Access Request"				'reviewing the wording for the email for any case that is not in a restricted basket and the resident is not on the phone.

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

STATS_manualtime = STATS_manualtime + (timer - start_time)						'This script allows for the writing of the email - so the manual time is adjusted as email length will vary
end_msg = "Thank you!" & vbCr & "Your request for access has been sent to QI Knowledge Now." & vbCr & vbCr
end_msg = end_msg & "Content of your Email to Knowledge Now:" & vbCr & "----------------------------------------------------------" & vbCr
end_msg = end_msg & "Subject: " & email_subject & vbCr & vbCr
end_msg = end_msg & email_body

call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/09/2021
'--Tab orders reviewed & confirmed----------------------------------------------11/09/2021
'--Mandatory fields all present & Reviewed--------------------------------------11/09/2021
'--All variables in dialog match mandatory fields-------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/09/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/09/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------11/09/2021
'--Script name reviewed---------------------------------------------------------11/09/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------11/09/2021
'--comment Code-----------------------------------------------------------------11/09/2021
'--Update Changelog for release/update------------------------------------------11/09/2021
'--Remove testing message boxes-------------------------------------------------11/09/2021
'--Remove testing code/unnecessary code-----------------------------------------11/09/2021
'--Review/update SharePoint instructions----------------------------------------11/09/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
