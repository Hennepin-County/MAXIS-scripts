'Required for statistical purposes==========================================================================================
name_of_script = "NOTES-NOMI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 276                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
call changelog_update("05/03/2016", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'logic to autofill the 'last_day_for_recert' field
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)
current_date = date


EMConnect ""
'grabs CAF date, turns CAF date into string for variable
CALL navigate_to_MAXIS_screen("STAT","PROG")

EMReadScreen err_msg, 7, 24, 02
IF err_msg = "BENEFIT" THEN	script_end_procedure ("Case must be in PEND II status for script to run, please update MAXIS panels TYPE & PROG (HCRE for HC) and run the script again.")

EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen fs_pend, 4, 10, 74

application_date = application_date & ""


IF cash_pend = "PEND" THEN
	EMReadScreen application_date, 8, 6, 33
	application_date = replace(application_date, " ", "/")
End if
IF cash2_pend = "PEND" THEN
	cash2_pend = CHECKED
ELSE
	cash2_pend = UNCHECKED
END If

IF cash1_pend = CHECKED OR cash2_pend = CHECKED THEN cash_pend = CHECKED


IF fs_pend = "PEND" THEN
	fs_pend = CHECKED
ELSE
	fs_pend = UNCHECKED
END IF

Call navigate_to_MAXIS_screen("CASE", "NOTE")
note_row = 5
day_before_app = DateAdd("d", -1, application_date) 'will set the date one day prior to app date'
Do
	EMReadScreen note_date, 8, note_row, 6
	EMReadScreen note_title, 55, note_row, 25
	note_title = trim(note_title)
	'MsgBox note_title
	IF note_title = "~ Appointment letter sent in MEMO for 5/4/2018 ~" then application_array(nomi_sent, case_entry) = note_date
	IF note_title = "~ Client missed application interview, NOMI sent via script ~" then application_array(nomi_sent, case_entry) = note_date
	IF note_title = "~ Appointment letter sent in MEMO ~" then application_array(appt_notc_sent, case_entry) = note_date
	IF left (note_title, 18) = "**New CAF received" then
		application_array(appt_notc_sent, case_entry) = note_date
		application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & "Wrong appt notice script used"
		END IF
		IF note_date = "        " then Exit Do
		note_row = note_row + 1
Loop until datevalue(note_date) < day_before_app

'creates interview date for 7 calendar days from the CAF date
interview_date = dateadd("d", 7, application_date)
If interview_date <= date then interview_date = dateadd("d", 7, date)
interview_date = interview_date & ""		'turns interview date into string for variable
'need to handle for if we dont need an appt letter, which would be...'

last_contact_day = CAF_date + 30
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

BeginDialog NOMI_dialog, 0, 0, 126, 95, "NOMI"
  EditBox 65, 5, 50, 15, MAXIS_case_number
  EditBox 65, 25, 50, 15, interview_date
  EditBox 65, 45, 50, 15, application_date
  ButtonGroup ButtonPressed
    OkButton 10, 70, 50, 15
    CancelButton 65, 70, 50, 15
  Text 10, 50, 50, 10, "Interview date:"
  Text 5, 30, 55, 10, "Application date:"
  Text 10, 10, 45, 10, "Case number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

DO
	DO
			Err_msg = ""
			Dialog NOMI_dialog	'dialog for all other users for ER
			cancel_confirmation
			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
			If isdate(interview_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date of needed interview."
			If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date of application."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
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
CALL start_a_new_spec_memo
	EMsendkey("************************************************************")
	Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
	Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & interview_date & ".")
	Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
	Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at ")
	Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & last_contact_day & " your application will be denied.") 'add 30 days
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to- face interview.")
	Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
	Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
	Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4
'Writes the case note for the NOMI
start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent via script ~ ")
	Call write_bullet_and_variable_in_case_note("* A notice was previously sent to client with detail about completing an interview. ")
	Call write_bullet_and_variable_in_case_note("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
	Call write_bullet_and_variable_in_case_note("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
	PF3
script_end_procedure("Success! The NOMI has been sent.")
