'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SELFPAY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog selfpay_dialog, 0, 0, 306, 105, "Self Pay"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 205, 5, 30, 15, dollar_amount1
  EditBox 255, 5, 45, 15, date1
  EditBox 110, 30, 30, 15, dollar_amount2
  DropListBox 155, 30, 80, 15, "Select one..."+chr(9)+"FMF"+chr(9)+"PSP"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
  EditBox 255, 30, 20, 15, number_of_days
  EditBox 195, 55, 45, 15, voucher_date_start
  EditBox 260, 55, 40, 15, voucher_date_end
  EditBox 55, 80, 135, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 195, 80, 50, 15
    CancelButton 250, 80, 50, 15
  Text 145, 35, 10, 10, "at"
  Text 140, 10, 65, 10, "Client will receive $"
  Text 280, 35, 25, 10, "nights."
  Text 10, 10, 45, 10, "Case number:"
  Text 240, 10, 10, 10, "on"
  Text 240, 35, 10, 10, "for"
  Text 10, 60, 180, 10, "Once self pay is verfied, client can be vouchered from:"
  Text 10, 35, 100, 10, "and has been told to self-pay $"
  Text 245, 60, 10, 10, "to"
  Text 10, 85, 40, 10, "Comments:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog selfpay_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If isnumeric(dollar_amount1) = false then err_msg = err_msg & vbNewLine & "* Enter a numeric dollar amount."		
		If date1 = "" then err_msg = err_msg & vbNewLine & "* Enter a date."
		If isnumeric(dollar_amount2) = False then err_msg = err_msg & vbNewLine & "* Enter a numeric dollar amount."
		If shelter_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the facility name"
		If number_of_days = "" then err_msg = err_msg & vbNewLine & "* Enter the number of days of stay."
		If voucher_date_start = "" then err_msg = err_msg & vbNewLine & "* Enter a voucher start date or 'n/a'."
		If voucher_date_end = "" then err_msg = err_msg & vbNewLine & "* Enter a voucher end date or 'n/a'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
		
'adding the case number 	
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### Self Pay ###")
Call write_variable_in_CASE_NOTE("* Client will receive $" & dollar_amount1 & " on " & date1 & ", and has been told to self-pay $" & dollar_amount2 & " at " & shelter_droplist & " Shelter for " & number_of_days & " nights.")
Call write_variable_in_CASE_NOTE("* Once self pay has been verfied, client can be vouchered from " & voucher_date_start & " to " & voucher_date_end)
Call write_variable_in_CASE_NOTE("* Self-Pay calculation agreement form given to client.")
Call write_variable_in_CASE_NOTE("* Shelter informed of need to self-pay")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")