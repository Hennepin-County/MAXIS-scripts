name_of_script = "UTILITIES - APPLICATION INQUIRY.vbs"
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
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("03/31/2022", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
function search_for_info_in_note(date_of_app, confirmation_number, name_of_applicant)
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

			case_note_header = trim(case_note_header)							'reformat the case note
			If case_note_date <> "        " Then case_note_date = DateAdd("d", 0, case_note_date)		'make the date a date

            'If it finds an application received case note for an online application
			If case_note_header = "~ Application Received (ApplyMN)" or _
                UCASE(case_note_header) = "~ APPLICATION RECEIVED (MNBENEFI" or _
                UCASE(case_note_header) = "~ APPLICATION RECEIVED (MN BENEF" then
				EmWriteScreen "X", note_row, 3									'open the case note
				transmit
                'In CASE/NOTE, reading 1st row and determining what is the date of application
                EMReadScreen case_note_header_row, 75, 4, 3
                case_note_header_row = trim(case_note_header_row)                   'trimming the header row
                IF instr(case_note_header_row, "for") THEN
                	length = len(case_note_header_row)                              'establishing the length of the variable
                	position = InStr(case_note_header_row, "for")                   'sets the position at the deliminator (in this case the 'for')
                	date_of_app = Right(case_note_header_row, length-position -3)   'establishes app date as three spaces after the deliminator
                    date_of_app = trim(replace(date_of_app, "~", ""))
                Else
                    date_of_app = ""                                                'defaulting to blank as a back-up
                End if

				the_row = 1														'search for the confirmation field
				the_col = 1
				EMSearch "* Confirmation #", the_row, the_col
				EMReadScreen confirmation_number, 30, the_row, 22		'Read the confirmation number
				confirmation_number = trim(confirmation_number)	'reformat
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

Call find_user_name(worker_name)						'defaulting the name of the user running the script
Call check_for_MAXIS(True)								'make sure we are in MAXIS
CALL MAXIS_case_number_finder(MAXIS_case_number)		'try to find the case number

'One and only dialog for this script
DO
    DO
        err_msg = ""
		Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 331, 185, "Application Inquiry"
            EditBox 140, 35, 50, 15, MAXIS_case_number
            CheckBox 10, 55, 225, 10, "Check here if this is for a NEW request - not known to MAXIS.", no_case_number_checkbox
            ButtonGroup ButtonPressed
            PushButton 100, 80, 135, 15, "Case Assignment SharePoint Page", case_assignment_btn
            EditBox 75, 115, 50, 15, date_of_app
            EditBox 230, 115, 75, 15, confirmation_number
            EditBox 65, 135, 90, 15, name_of_applicant
            PushButton 170, 135, 135, 15, "Read application info from CASE NOTE", collect_from_case_note_btn
            EditBox 65, 160, 145, 15, worker_name
            OkButton 215, 160, 50, 15
            CancelButton 270, 160, 50, 15
            PushButton 125, 115, 15, 15, "?", app_date_question
            PushButton 305, 115, 15, 15, "?", confirmation_number_question
            PushButton 155, 135, 15, 15, "?", applicant_name_question
            PushButton 305, 135, 15, 15, "?", read_note_question
            GroupBox 5, 70, 320, 30, "APPL Information"
            Text 10, 120, 65, 10, "Date of Application:"
            Text 10, 165, 55, 10, "Sign your Email"
            Text 10, 40, 130, 10, "Case Number with the lost application:"
            GroupBox 5, 105, 320, 50, "Application Detail"
            Text 10, 140, 55, 10, "Applicant Name: "
            Text 155, 120, 75, 10, "Confirmation Number:"
            Text 10, 10, 315, 20, "If a resident is reporting they have submitted an online application, and there is no coresponding application in ECF, this script can assist in sending the request to QI to locate."
        EndDialog

        Dialog Dialog1
        cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		If no_case_number_checkbox = checked then err_msg = ""			'if the checkbox is checked it will blank out the case number error messaging
        If (no_case_number_checkbox = checked and trim(MAXIS_case_number) <> "") then err_msg = err_msg & vbNewLine & "* Enter either the case number or check the box to confirm no case number exists, not both options."
        If (no_case_number_checkbox = unchecked and trim(MAXIS_case_number) = "") then err_msg = err_msg & vbNewLine & "* Enter either the case number or check the box to confirm no case number exists, not both options"
        If IsDate(date_of_app) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of application as a valid date."
		confirmation_number = trim(confirmation_number)
		If confirmation_number = "" Then err_msg = err_msg & vbNewLine & "* Enter the confirmation number for the online application."
		name_of_applicant = trim(name_of_applicant)
		If name_of_applicant = "" Then err_msg = err_msg & vbNewLine & "* Enter the name of the person who submitted the online application."
		worker_name = trim(worker_name)
		If worker_name = "" Then err_msg = err_msg & vbNewLine & "* Enter your name to sign your email."

		If ButtonPressed = collect_from_case_note_btn Then
			err_msg = "LOOP" & err_msg
			Call search_for_info_in_note(date_of_app, confirmation_number, name_of_applicant)		'This will call the in script function to fill in some of the fields
		End If

        'Opening Case Assignment SharePoint Online Page
        If ButtonPressed = case_assignment_btn Then
			err_msg = "LOOP" & err_msg
			Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Case%20Assignment%20(CA)%20Team.aspx")
		End If

		If ButtonPressed = read_note_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("When can you use the functionality to read online application Information from CASE:NOTE and what will it read?"& vbNewLine & vbNewLine & "*** THE NOTE THIS REFERS TO IS THE NOTES - APPLICATION RECEIVED NOTE WITH ONLINE APPLICATION INFORMATION. ****" & vbNewLine & vbNewLine & " - If the case was pended and the script NOTES - Application Received is used to capture the information then the details are listed in CASE:NOTE." & vbNewLine & " - What can be read:" & vbNewLine & "    *Application Date" & vbNewLine & "    *Confirmation Number" & vbNewLine & "    *Name of Applicant", vbInformation, "What can e read from CASE:OTE")
		End If

		If ButtonPressed = confirmation_number_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("The confirmation number is the primary means to search through the files of MNbenefit applications submitted" & vbNewLine & vbNewLine & "There is no other information that allows us to search as quickly for MNbenefit applications." & vbNewLine & vbNewLine & "--- What if the client does not know their Confirmation Number?" & vbNewLine & "Information from when the client submitted the application online can be found from logging back into their account for online application.", vbInformation, "Why do we need the confirmation number?")
		End If

		If ButtonPressed = applicant_name_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("The application name allows us to most easily confirm the application is correct as it is unique identifying information. Enter this information as it was entered in the application if possible.", vbInformation, "Why is the applicant name needed?")
		End If

		If ButtonPressed = app_date_question Then
			err_msg = "LOOP" & err_msg
			info_msg = MsgBox("There are hundreds of online application applications submitted in any day. This number has increased more recently." & vbNewLine & vbNewLine & "*** The date needs to be the precise date the application was submitted ***" & vbNewLine & vbNewLine & "This is important because the appplications must be pulled by date before searching. Trying to pull either the wrong date or mutliple dates prevents us from finding this information." & vbNewLine & vbNewLine & "Information from when the client submitted the application online can be found from logging back into their account for online application.", vbInformation, "Why is the application date needed?")
		End If

		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------Email
email_subject = "Request for Search for Online Application not in ECF"			'Setting the subject line
email_body = "Please recover the Online Application file that cannot be found." & vbCr & vbCr						'Fillin i nthe information for the body of the email

If no_case_number_checkbox = unchecked Then email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr
If no_case_number_checkbox = checked Then email_body = email_body & "No Case Number known at this time as case has not been pended." & vbCr
email_body = email_body & "Name of Applicant: " & name_of_applicant & vbCr
email_body = email_body & "Date of Application: " & date_of_app & vbCr
email_body = email_body & "Confirmation Number: " & confirmation_number & vbCr

'using case pending or not to give information to QI
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
If case_pending = TRUE Then email_body = email_body & vbCr & "~~The case has been APPL'd. ~~" & vbCr
If case_pending = FALSE Then email_body = email_body & vbCr & "~~This case is not pending and has not been APPL'd.~~" & vbCr

email_body = email_body & vbCr & "Thank you, " & vbCr & worker_name
Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)		'Send the Email

'Add a message with the email information for display
end_msg = "Success!" & vbNewLine & vbNewLine
end_msg = end_msg & "Your email has been sent to QI. Email sent:" & vbCr
end_msg = end_msg & "--------------------------------------------------" & vbCr & vbCr
end_msg = end_msg & email_body

call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/31/2022
'--Tab orders reviewed & confirmed----------------------------------------------03/31/2022
'--Mandatory fields all present & Reviewed--------------------------------------03/31/2022
'--All variables in dialog match mandatory fields-------------------------------03/31/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------03/31/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------03/31/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------03/31/2022-------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/31/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------03/31/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------03/31/2022-------------------N/A
'--Out-of-County handling reviewed----------------------------------------------03/31/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/31/2022
'--BULK - review output of statistics and run time/count (if applicable)--------03/31/2022-------------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------03/31/2022
'--Incrementors reviewed (if necessary)-----------------------------------------03/31/2022-------------------N/A
'--Denomination reviewed -------------------------------------------------------03/31/2022
'--Script name reviewed---------------------------------------------------------03/31/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------03/31/2022-------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete-----------------------------------------
'--comment Code-----------------------------------------------------------------03/31/2022
'--Update Changelog for release/update------------------------------------------03/31/2022
'--Remove testing message boxes-------------------------------------------------03/31/2022
'--Remove testing code/unnecessary code-----------------------------------------03/31/2022
'--Review/update SharePoint instructions----------------------------------------04/01/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/01/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/01/2022
'--Complete misc. documentation (if applicable)---------------------------------03/31/2022
'--Update project team/issue contact (if applicable)----------------------------03/31/2022
