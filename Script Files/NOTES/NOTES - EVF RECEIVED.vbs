'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EVF RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 120          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'main dialog
BeginDialog EVF_received, 0, 0, 291, 205, "Employment Verification Form Received"
  EditBox 70, 5, 60, 15, MAXIS_case_number
  EditBox 220, 5, 60, 15, date_received
  ComboBox 70, 30, 210, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
  EditBox 70, 50, 210, 15, employer
  EditBox 70, 70, 210, 15, client
  DropListBox 75, 110, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", info
  EditBox 220, 110, 60, 15, info_date
  EditBox 75, 130, 60, 15, request_info
  CheckBox 160, 135, 105, 10, "10 day TIKL for additional info", Tikl_checkbox
  EditBox 70, 160, 210, 15, actions_taken
  EditBox 70, 180, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 180, 50, 15
    CancelButton 230, 180, 50, 15
  Text 10, 135, 65, 10, "Info Requested via:"
  Text 10, 115, 60, 10, "Addt'l Info Reqstd:"
  Text 5, 75, 60, 10, "Household Memb:"
  Text 10, 55, 55, 10, "Employer name:"
  Text 15, 165, 50, 10, "Actions taken:"
  Text 5, 185, 60, 10, "Worker Signature:"
  Text 25, 35, 40, 10, "EVF Status:"
  Text 150, 10, 65, 10, "Date EVF received:"
  Text 15, 10, 50, 10, "Case Number:"
  Text 160, 115, 55, 10, "Date Requested:"
  GroupBox 5, 95, 280, 60, "Is additional information needed?"
EndDialog

'The script----------------------------------------------------------------------------------------------------
'connects to BlueZone & grabs the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Tikl_checkbox = checked 'defaulting the TIKL checkbox to be checked initially in the dialog. 

'starts the EVF received case note dialog
DO
	Do
		err_msg = ""
		Dialog EVF_received       	'starts the EVF dialog
		cancel_confirmation 		'asks if you want to cancel and if "yes" is selected sends StopScript
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You need to type a valid case number."
		IF IsDate(date_received) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date for date the EVF was received."
		If EVF_status_dropdown = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select the status of the EVF on the dropdown menu"		'checks that there is a date in the date received box
		IF employer = "" THEN err_msg = err_msg & vbCr & "* You must enter the employers name."  'checks if the employer name has been entered
		IF client = "" THEN err_msg = err_msg & vbCr & "* You must enter the MEMB information."  'checks if the client name has been entered
		IF info = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select if additional info was requested."  'checks if completed by employer was selected
		IF info = "yes" and IsDate(info_date) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date that additional info was requested."  'checks that there is a info request date entered if the it was requested
		IF info = "yes" and request_info = "" THEN err_msg = err_msg & vbCr & "* You must enter the method used to request additional info."		'checks that there is a method of inquiry entered if additional info was requested		
		If info = "no" and request_info <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have information requested."	
		If info = "no" and info_date <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have a date requested."	
		If Tikl_checkbox = 1 and info <> "yes" then err_msg = err_msg & vbCr & "* Additional informaiton was not requested, uncheck the TIKL checkbox."	
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered		
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!" 		'checks that the case note was signed
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "* Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false			

'starts a blank case note
call start_a_blank_case_note
'this enters the actual case note info 
call write_variable_in_CASE_NOTE("***EVF received " & date_received & ": " & EVF_status_dropdown & "***")
Call write_bullet_and_variable_in_CASE_NOTE("Employer name", employer)
Call write_bullet_and_variable_in_CASE_NOTE("EVF for HH member", client)
'for additional information needed
IF info = "yes" then 
	call write_variable_in_CASE_NOTE ("* Additional Info requested: " & info & " on " & info_date & " by " & request_info)
	If Tikl_checkbox = 1 then call write_variable_in_CASE_NOTE ("***TIKLed for 10 day return.***")
Else 
	Call write_variable_in_CASE_NOTE("* No additional information is needed/requested.")
END IF 
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Checks if additional info is yes and the TIKL is checked, sets a TIKL for the return of the info
IF Tikl_checkbox = 1 THEN 
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18)		'The following will generate a TIKL formatted date for 10 days from now.
	call write_variable_in_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action. (TIKL auto-generated from script)." )
	transmit
	PF3
	'Success message
	script_end_procedure("Success! TIKL has been sent for 10 days from now for the additional information requested.")
ELSE 
	script_end_procedure("")	'ends the script without the success message
End if
