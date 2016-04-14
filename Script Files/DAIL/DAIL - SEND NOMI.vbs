'STATS gathering
name_of_script = "DAIL - SEND NOMI.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 276         'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

'Dialogs----------------------------------------------------------------------------------------------------
BeginDialog Hennepin_worker_signature, 0, 0, 186, 100, "Hennepin County worker signature and client region"
  DropListBox 80, 10, 100, 15, "Select one..."+chr(9)+"Central/NE"+chr(9)+"North"+chr(9)+"Northwest"+chr(9)+"South MPLS"+chr(9)+"S. Suburban"+chr(9)+"West", region_residence
  EditBox 80, 30, 55, 15, last_day_for_recert
  EditBox 80, 50, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 75, 50, 15
    CancelButton 130, 75, 50, 15
  Text 10, 55, 60, 10, "Worker signature:"
  Text 5, 15, 70, 10, "Region of residence: "
  Text 10, 35, 65, 10, "Last day for recert:"
EndDialog

BeginDialog worker_signature_dialog, 0, 0, 191, 80, "Worker signature"
  EditBox 80, 10, 55, 15, last_day_for_recert
  EditBox 80, 30, 105, 15, worker_signature
  ButtonGroup ButtonPressed_worker_signature_dialog
    OkButton 80, 50, 50, 15
    CancelButton 135, 50, 50, 15
  Text 5, 35, 70, 10, "Sign your case note:"
  Text 10, 15, 65, 10, "Last day for recert:"
EndDialog

'logic to autofill the 'last_day_for_recert' into the notice
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_for_recert = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string			

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER; As such, it does NOT include protections to be ran independently.
EMConnect ""
EMSendKey "x"
transmit

'Reading date and time of recertification appointment from the TIKL--DAIL message that should be read is: "~*~*~CLIENT HAD RECERT INTERVIEW APPT AT..." This is the part that is static in the DAIL message
EMReadScreen interview_date_time, 19, 9, 46    'reads "MM/DD/YYYY HH:MM PM" (or any combination less) off of dail messate
row  = 1
col = 1
EMSearch "Case Number: ", row, col
If row =- 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
EMReadScreen case_number, 8, row, col + 12
case_number = trim(case_number)
PF3 			'removes the TIKL window
'navigates to CASE/NOTE to user can see if interview has been completed or not
EMSendKey "n"
transmit

'Msgbox asking the user misssed their interview
interview_confirm = MsgBox("Was an interview completed for this case's recertification?", vbYesNoCancel, "Interview confirmation")
	If interview_confirm = vbCancel then stopscript
	If interview_confirm = vbYes then interview_confirm = TRUE 
	If interview_confirm = vbNo then interview_confirm = FALSE

If interview_confirm = TRUE then 
	PF3 	'returns user back to DAIL/DAIL and stops the script since no further action is required
	script_end_procedure("Success! A NOMI is not required if the recertification interview is complete." & vbNewLine & "Please review the case for completion if necessary.")
ELSE
	'Msgbox asking the user to confirm if the client has sent a CAF or if no contact has been made by the client
	recert_forms_confirm = MsgBox("A NOMI is NOT needed when a SNAP recipient has not made contact with the agency about their recertification, AND the CAF has not been received." & vbNewLine & "See TE.02.05.15" & vbNewLine & vbNewLine & "Press Yes to send the NOMI." & _
	vbNewLine & "Press No if client contact has not been made with the agency." & vbNewLine & "Press Cancel to end the script.", vbYesNoCancel, "Client contact confirmation")
		If recert_forms_confirm = vbCancel then stopscript
		If recert_forms_confirm = vbYes then result_of_msgbox = TRUE
		If recert_forms_confirm = vbNo then result_of_msgbox = FALSE
END IF

If result_of_msgbox = FALSE then		'if false a case note will be made, but a NOMI will not be sent as this is not necessary. 
	dialog worker_signature_dialog
	If ButtonPressed_worker_signature_dialog = 0 then stopscript
	Call start_a_blank_case_note 'Navigates to a blank case note & writes the case note
	Call write_variable_in_CASE_NOTE ("**Client missed SNAP recertification interview**")
	Call write_variable_in_CASE_NOTE("* Interview appointment was scheduled for: " & interview_date_time)
	Call write_variable_in_CASE_NOTE ("* A SNAP NOMI for recertification SPEC/MEMO has not been sent.")
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE (worker_signature & ", using automated script.")
	PF3	'saves the case note'
	Call navigate_to_MAXIS_screen("DAIL", "DAIL")	'brings user back to DAIL'
	script_end_procedure("A SNAP NOMI for recertification case note has been made, but a SPEC/MEMO has NOT been sent." & vbNewLine & vbNewLine & _
	"Per POLI/TEMP TE02.05.15: When there is no request for further assistance the client will receive the proper closing (the autoclose notice).")
END IF

IF result_of_msgbox = TRUE then		'user pressed YES button, SPEC/MEMO will be sent
	If worker_county_code = "x127" then
		dialog Hennepin_worker_signature		'dialog for Hennepin users with county office selection options
		Else
		dialog worker_signature_dialog			'dialog for everyone else...because elitism:) 
		End if
	If ButtonPressed_worker_signature_dialog = 0 then stopscript
	PF3							'exits case note, back to DAIL
	EMSendKey "p"				'navigates to SPEC
	transmit
	EMWriteScreen "MEMO", 20, 70		'navigates to MEMO'
	transmit
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	
	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process

	If worker_county_code = "x127" then
		'writes in the SPEC/MEMO for Hennepin County users
		Call write_variable_in_SPEC_MEMO("************************************************************")
		Call write_variable_in_SPEC_MEMO("You have missed your SNAP interview that was scheduled for " & interview_date_time)
		Call write_variable_in_SPEC_MEMO(" ")
	  Call write_variable_in_SPEC_MEMO("Please contact your worker at 612-596-1300 to complete the required SNAP interview.")
		IF region_residence = "Central/NE" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Human Services building office to complete an interview. The office is located at: 525 Portland Ave. in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "North" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the North Minneapolis office to complete an interview. The office is located at: 1001 Plymouth Ave. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
	  ELSEIF region_residence = "Northwest" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Brooklyn Center to complete an interview. The office is located at: 7051 Brooklyn Blvd. Office hours are Monday through Friday from 7:30 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "South MPLS" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Century Plaza office to complete an interview. The office is located at: 330 S. 12th Street in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "S. Suburban" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Bloomington office complete an interview. The office is located at: 9600 Aldrich Ave. Office hours are Monday, Tuesday, Wednesday and Friday from 8 a.m. to 4:30 p.m. and Thursday from 8 a.m. to 6:30 p.m.")
		ElseIF region_residence = "West" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Hopkins office to complete an interview. The office is located at: 1011 1st Street S. (in the Wells Fargo building). Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		END IF
		Call write_variable_in_SPEC_MEMO(" ")
	  Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your renewal must be completed by " & last_day_for_recert & ", or your SNAP case will Auto-Close on this date.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
	ELSE
		'Writes the info into the SPEC/MEMO for other users
		Call write_variable_in_SPEC_MEMO("************************************************************")
		Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & interview_date_time)
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required Food Support interview.")
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
	END IF
	PF4	'saves and exits from SPEC/MEMO
	PF3
	Call start_a_blank_case_note 'Navigates to a blank case note & writes the case note
	Call write_variable_in_CASE_NOTE ("**Client missed SNAP recertification interview**")
	Call write_variable_in_CASE_NOTE("* Appointment was scheduled for: " & interview_date_time)
	Call write_variable_in_CASE_NOTE ("* A SNAP NOMI for recertification SPEC/MEMO has been sent.")
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE (worker_signature & ", using automated script.")
	PF3	'saves the case note'
	Call navigate_to_MAXIS_screen("DAIL", "DAIL") 'brings user back to DAIL'
	script_end_procedure("Success! A SNAP NOMI for recertification SPEC/MEMO has been sent.")	
END IF
