'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - APPOINTMENT LETTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 195                               'manual run time in seconds
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
call changelog_update("06/22/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call MAXIS_case_number_finder (MAXIS_case_number)

BeginDialog case_number_dlg, 0, 0, 131, 45, "Case Number"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 55, 25, 35, 15
    CancelButton 95, 25, 30, 15
  Text 10, 10, 50, 10, "Case Number:"
EndDialog

Do
    err_msg = ""

    Dialog case_number_dlg
    If buttonpressed = Cancel Then script_end_procedure("")

    If len(MAXIS_case_number) >8 Then err_msg = err_msg & vbNewLine & "* Case numbers should not be more than 8 numbers long."
    If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* Check the case number, it appears to be invalid."

    If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

'grabs CAF date, turns CAF date into string for variable
call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)
application_date = application_date & ""

'creates interview date for 7 calendar days from the CAF date
interview_date = dateadd("d", 7, application_date)
If interview_date <= date then interview_date = dateadd("d", 7, date)
interview_date = interview_date & ""		'turns interview date into string for variable

BeginDialog appt_dialog, 0, 0, 121, 75, "APPOINTMENT LETTER"
  EditBox 65, 5, 50, 15, application_date
  EditBox 65, 25, 50, 15, interview_date
  ButtonGroup ButtonPressed
    OkButton 10, 50, 50, 15
    CancelButton 65, 50, 50, 15
  Text 10, 30, 50, 10, "Interview date:"
  Text 5, 10, 55, 10, "Application date:"
EndDialog



 'need to handle for if we dont need an appt letter, which would be...'

Do
	Do
		err_msg = ""
		dialog appt_dialog
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
call start_a_new_spec_memo                                                   'Transmits to start the memo writing process
EMsendkey("************************************************************")
Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & ".")
Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("The interview must be completed by " & interview_date & ".")
Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
Call write_variable_in_SPEC_MEMO("You may be able to have SNAP benefits issued within 24 hours of the interview.")
'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")
Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " your application will be denied.") 'add 30 days
Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3).")
'Call write_variable_in_SPEC_MEMO("************************************************************")
PF4
'msgbox "should be all memoed out"

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
Call write_variable_in_CASE_NOTE("A notice has been sent via SPEC/MEMO informing the client of needed interview.")
Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
PF3
script_end_procedure("Success! The Appointment Letter has been sent.")
'IF action_completed = False then 'build handling'
'    script_end_procedure ("Warning! Appointment letter was not sent. Check the case manually.")
'Else
'    script_end_procedure ("Case has been updated please review to ensure it was processed correctly.")
'End if
