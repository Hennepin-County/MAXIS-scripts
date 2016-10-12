'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - COUNTY BURIAL DETERMINATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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



'Script---------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Call navigate_to_MAXIS_screen ("CASE", "NOTE")
row = 1
col = 1
EMSearch "County Burial Application", row, col 
IF row <> 0 THEN 
	application_note = TRUE 
	EMWriteScreen "x", row, 3
	transmit
ELSEIF row = 0 THEN 
	application_note = FALSE
END IF 

'Script is gathering the service information from the application case note
IF application_note = TRUE THEN
	row = 1
	col = 1 
	EMSearch "Services Requested:", row, col 		'Looking for what type of service requested (burial or cremation)
	If row <> 0 Then 
		EMReadScreen service_in_note, 55, row, 25	'Reading the service requested from the case note
		service_in_note = trim(service_in_note)
		
		services_dropdown = service_in_note+chr(9)+"Creamation"+chr(9)+"Burial"	'Formatting the dropdown for the next dialog
	End If 
	
	row = 1
	col = 1 
	EMSearch "Amount requested:", row, col 			'getting the information from the case note with the amount requested
	EMReadScreen service_cost, 7, row, col+18
	service_cost = trim(service_cost)
End IF 

If services_dropdown = "" Then services_dropdown = ""+chr(9)+"Cremation"+chr(9)+"Burial"	'If a note was not found - creating the dropdown

'Dialog-----------------Defined here so dropdown can be dynamic-----------------------------------------
BeginDialog county_burial_determination_dialog, 0, 0, 281, 130, "County Burial Determination"
  EditBox 55, 5, 50, 15, MAXIS_case_number
  ComboBox 185, 5, 90, 45, ""+chr(9)+"Approved"+chr(9)+"Denied", burial_request_status
  EditBox 65, 25, 40, 15, service_cost
  ComboBox 185, 25, 90, 45, services_dropdown, requested_services
  EditBox 5, 55, 270, 15, denial_reason
  EditBox 5, 85, 270, 15, other_notes
  EditBox 70, 110, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 170, 110, 50, 15
    CancelButton 225, 110, 50, 15
  Text 5, 10, 50, 10, "Case Number"
  Text 115, 10, 65, 10, "Burial Request was"
  Text 5, 30, 55, 10, "Cost of services"
  Text 115, 30, 70, 10, "Services requested"
  Text 5, 45, 65, 10, "Reason for Denial"
  Text 5, 75, 25, 10, "Notes"
  Text 10, 115, 60, 10, "Worker Signature"
EndDialog

'Running the dialog
Do
	Do 
		err_msg = ""
		Dialog county_burial_determination_dialog
		cancel_confirmation
		If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter the case number."
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "You must sign your case note."
		If burial_request_status = "" Then err_msg = err_msg & vbNewLine & "Please select or fill in the decision made on this case."
		If requested_services = "" Then err_msg = err_msg & vbNewLine & "Enter the service requested on the application."
		If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***County Burial " & burial_request_status & "***")
CALL write_bullet_and_variable_in_CASE_NOTE("Service requested", requested_services)
CALL write_bullet_and_variable_in_CASE_NOTE("Amount", service_cost)
CALL write_bullet_and_variable_in_CASE_NOTE("Denied due to", denial_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Success! Case note completed.")
