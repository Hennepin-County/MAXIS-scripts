'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - NOMI.vbs"
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/18/2018", "Updated for On Demand Waiver processing.", "MiKayla Handley, Hennepin County")
call changelog_update("01/10/2017", "Updated TIKL functionality. A TIKL is created for Application Day 30 if NOMI is sent prior to Application Day 30. Otherwise a TIKL is created for an additional 10 days .", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Resolved merge conflict error.", "Ilse Ferris, Hennepin County")
call changelog_update("11/21/2016", "Removed Hennepin County specific NOMI process. Users will follow the process documented in POLI/TEMP TE02.05.15. Added TIKL to follow up on the application's progress. Added intial case number dialog to allow for the application date to be autofilled into the NOMI dialog. Removed message box to identify if case is a renewal. Replaced with a check box on the initial dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("11/20/2016", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'logic to autofill the 'last_day_for_recert' field
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_for_recert = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)
application_date = application_date & ""
'creates interview date for 7 calendar days from the CAF date
interview_date = dateadd("d", 7, application_date)
If interview_date <= date then interview_date = dateadd("d", 7, date)
interview_date = interview_date & ""		'turns interview date into string for variable
'need to handle for if we dont need an appt letter, which would be...'

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog NOMI_dialog, 0, 0, 126, 95, "NOMI"
  EditBox 65, 5, 50, 15, MAXIS_case_number
  EditBox 65, 25, 50, 15, application_date
  EditBox 65, 45, 50, 15, interview_date
  ButtonGroup ButtonPressed
    OkButton 10, 70, 50, 15
    CancelButton 65, 70, 50, 15
  Text 10, 50, 50, 10, "Interview date:"
  Text 5, 30, 55, 10, "Application date:"
  Text 10, 10, 45, 10, "Case number:"
EndDialog
Do
	Do
		err_msg = ""
		dialog NOMI_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	Loop until err_msg = ""
call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

	last_contact_day = dateadd("d", 30, application_date)
	If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date
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
    PF3
  'Writes the case note for the NOMI
  start_a_blank_CASE_NOTE
  	Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent via script ~ ")
  	Call write_variable_in_CASE_NOTE("A notice was previously sent to client with detail about completing an interview. ")
  	Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
    Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
  	Call write_variable_in_CASE_NOTE("---")
  	Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
  	PF3
  script_end_procedure("Success! The NOMI has been sent.")
