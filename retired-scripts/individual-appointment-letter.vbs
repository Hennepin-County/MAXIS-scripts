'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - APPOINTMENT LETTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 350                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("05/01/2022", "Updated the Appointment Notice to have information for residents about in person support.", "Casey Love, Hennepin County")
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
call changelog_update("08/01/2021", "Changed the notices sent in 2 ways:##~## ##~## - Updated verbiage on how to submit documents to Hennepin.##~## ##~## - Appointment Notices will now be sent with a date of 5 days from the date of application.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2020", "Update to the notice wording. Information and direction for in-person interview option removed. County offices are not currently open due to the COVID-19 Peacetime Emergency.", "Casey Love, Hennepin County")
call changelog_update("07/20/2018", "Updated verbiage of notice and changed interview date to be 10 days from app date.", "Casey Love, Hennepin County")
call changelog_update("06/22/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call check_for_MAXIS(TRUE)
Call MAXIS_case_number_finder (MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 65, "Case number"
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 70, 25, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 45, 50, 15
    CancelButton 125, 45, 45, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 60, 10, "Worker signature:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'grabs CAF date, turns CAF date into string for variable
call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)

interview_date = dateadd("d", 5, application_date)
If interview_date <= date then interview_date = dateadd("d", 5, date)

Call change_date_to_soonest_working_day(interview_date, "FORWARD")

application_date = application_date & ""
interview_date = interview_date & ""

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 266, 80, "APPOINTMENT LETTER"
EditBox 185, 20, 55, 15, interview_date
ButtonGroup ButtonPressed
	OkButton 155, 60, 50, 15
	CancelButton 210, 60, 50, 15
EditBox 50, 20, 55, 15, application_date
Text 120, 25, 60, 10, "Appointment date:"
GroupBox 5, 5, 255, 35, "Enter a new appointment date only if it's a date county offices are not open."
Text 15, 25, 35, 10, "CAF date:"
Text 25, 45, 205, 10, "If same-day interview is being offered please use today's date"
EndDialog
 'need to handle for if we dont need an appt letter, which would be...'

Do
	Do
		err_msg = ""
		dialog Dialog1
		cancel_confirmation
        If isdate(application_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid application date."
		If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid interview date."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Figuring out the last contact day
last_contact_day = dateadd("d", 30, application_date)
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

'This checks to make sure the case is not in background and is in the correct footer month for PND1 cases.
Do
	call navigate_to_MAXIS_screen("STAT", "SUMM")
	EMReadScreen month_check, 11, 24, 56 'checking for the error message when PND1 cases are not in APPL month
	IF left(month_check, 5) = "CASES" THEN 'this means the case can't get into stat in current month
		EMWriteScreen mid(month_check, 7, 2), 20, 43 'writing the correct footer month (taken from the error message)
		EMWriteScreen mid(month_check, 10, 2), 20, 46 'writing footer year
		EMWriteScreen "STAT", 16, 43
		EMWriteScreen "SUMM", 21, 70
		transmit 'This transmit should take us to STAT / SUMM now
	END IF
	'This section makes sure the case isn't locked by background, if it is it will loop and try again
	EMReadScreen SELF_check, 4, 2, 50
	If SELF_check = "SELF" then
		PF3
		Pause 2
	End if
Loop until SELF_check <> "SELF"

'Navigating to SPEC/MEMO
call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)                                                 'Transmits to start the memo writing process

Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & interview_date & ". **")
Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
Call write_variable_in_SPEC_MEMO(" ")
CALL write_variable_in_SPEC_MEMO("All interviews are completed via phone. If you do not have a phone, go to one of our Digital Access Spaces at any Hennepin County Library or Service Center. No processing, no interviews are completed at these sites. Some Options:")
CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
CALL write_variable_in_SPEC_MEMO(" ")
CALL digital_experience
CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can always request a paper copy via phone.")


PF4
'msgbox "should be all memoed out"

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
Call write_variable_in_CASE_NOTE("A notice has been sent via SPEC/MEMO informing the client of needed interview.")
Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
PF3
script_end_procedure_with_error_report("Success! The Appointment Letter has been sent.")
'IF action_completed = False then 'build handling'
'    script_end_procedure ("Warning! Appointment letter was not sent. Check the case manually.")
'Else
'    script_end_procedure ("Case has been updated please review to ensure it was processed correctly.")
'End if
