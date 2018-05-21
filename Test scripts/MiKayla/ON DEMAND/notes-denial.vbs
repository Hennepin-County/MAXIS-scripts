'Required for statistical purposes==========================================================================================
name_of_script = "notes-denial.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 276                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
CALL changelog_update("05/03/2018", "Initial version.", "MiKayla Handley, Hennepin County")
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

CALL ONLY_create_MAXIS_friendly_date(current_date)
EMConnect ""
'grabs CAF date, turns CAF date into string for variable
CALL autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)
application_date = application_date & ""

'creates interview date for 7 calendar days from the CAF date
interview_date = dateadd("d", 7, application_date)
If interview_date <= date then interview_date = dateadd("d", 7, date)
interview_date = interview_date & ""		'turns interview date into string for variable
'need to handle for if we dont need an appt letter, which would be...'

last_contact_day = CAF_date + 30
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

BeginDialog denial_dialog, 0, 0, 241, 65, "PND2 DENIAL"
  EditBox 60, 20, 50, 15, MAXIS_case_number
  EditBox 180, 20, 50, 15, interview_date
  ButtonGroup ButtonPressed
    OkButton 125, 45, 50, 15
    CancelButton 180, 45, 50, 15
  Text 10, 25, 45, 10, "Case number:"
  Text 120, 25, 55, 10, "Application date:"
  GroupBox 5, 5, 230, 35, "ONLY use for denials due to no interview on REPT/PND2"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

DO
	DO
			Err_msg = ""
			Dialog denial_dialog	'dialog for all other users for ER
			cancel_confirmation
			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
			If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date of application."			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
	'Figuring out the last contact day

  Do
  	CALL navigate_to_MAXIS_screen("STAT", "SUMM")
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
  	CALL write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
  	CALL write_variable_in_SPEC_MEMO("Your interview was not completed by " & last_contact_day & " and your application has been denied.") 'add 30 days
  	CALL write_variable_in_SPEC_MEMO(" ")
  	CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
  	CALL write_variable_in_SPEC_MEMO("You can also request a paper copy.")
  	CALL write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
  	CALL write_variable_in_SPEC_MEMO("************************************************************")
  	PF4

  'Writes the case note for the NOMI
  start_a_blank_CASE_NOTE
  	CALL write_variable_in_CASE_NOTE("~ Client has not completed application interview, DENIAL via script ~ ")
  	CALL write_bullet_and_variable_in_case_note("A notice was previously sent to client with detail about completing an interview. ")
  	CALL write_bullet_and_variable_in_case_note("Application date", application_date)
  	CALL write_bullet_and_variable_in_case_note("NOMI sent date", interview_date)
  	CALL write_variable_in_CASE_NOTE("---")
  	CALL write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
  	PF3

  script_end_procedure("Please snsure that the corrects steps are taken when denying off REPT/PND2.")
