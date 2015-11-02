'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EXPEDITED SCREENING.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog exp_screening_dialog, 0, 0, 181, 210, "Expedited Screening Dialog"
  EditBox 55, 5, 95, 15, case_number
  EditBox 100, 25, 50, 15, income
  EditBox 100, 45, 50, 15, assets
  EditBox 100, 65, 50, 15, rent
  CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 75, 95, 45, 10, "Electricity", electric_check
  CheckBox 130, 95, 35, 10, "Phone", phone_check
  DropListBox 70, 115, 105, 15, "intake"+chr(9)+"add-a-program", application_type
  EditBox 70, 135, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 155, 50, 15
    CancelButton 125, 155, 50, 15
  Text 10, 180, 160, 15, "The income, assets and shelter costs fields will default to $0 if left blank. "
  Text 5, 30, 95, 10, "Income received this month:"
  Text 5, 50, 95, 10, "Cash, checking, or savings: "
  Text 5, 70, 90, 10, "AMT paid for rent/mortgage:"
  GroupBox 5, 85, 170, 25, "Utilities claimed (check below):"
  Text 5, 120, 60, 10, "Application is for:"
  Text 5, 140, 60, 10, "Worker signature:"
  Text 5, 10, 50, 10, "Case number: "
  GroupBox 0, 170, 175, 30, "**IMPORTANT**"
EndDialog

'DATE BASED LOGIC FOR UTILITY AMOUNTS------------------------------------------------------------------------------------------
If date >= cdate("10/01/2015") then			'these variables need to change in October 2014, and subsequently every October
	heat_AC_amt = 454
	electric_amt = 141
	phone_amt = 38
Else
	heat_AC_amt = 450
	electric_amt = 150
	phone_amt = 38
End if

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""
'It will search for a case number.
call MAXIS_case_number_finder(case_number)

'Shows the dialog
Do
	Do
		Do
			Dialog exp_screening_dialog
			cancel_confirmation
			If isnumeric(case_number) = False then MsgBox "You must enter a valid case number."
		Loop until isnumeric(case_number) = True
		If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) then MsgBox "The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
	Loop until (income = "" or isnumeric(income) = True) and (assets = "" or isnumeric(assets) = True) and(rent = "" or isnumeric(rent) = True)
	If worker_signature = "" then MsgBox "You must sign your case note."
Loop until worker_signature <> ""

'checking for an active MAXIS session
Call check_for_MAXIS(FALSE)

'LOGIC AND CALCULATIONS----------------------------------------------------------------------------------------------------
'Logic for figuring out utils. The highest priority for the if...then is heat/AC, followed by electric and phone, followed by phone and electric separately.
If heat_AC_check = checked then
	utilities = heat_AC_amt
ElseIf electric_check = checked and phone_check = checked then 
	utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
ElseIf phone_check = checked and electric_check = unchecked then
	utilities = phone_amt
ElseIf electric_check = checked and phone_check = unchecked then
	utilities = electric_amt
End if

'in case no options are clicked, utilities are set to zero.
If phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked then utilities = 0

'If nothing is written for income/assets/rent info, we set to zero.
If income = "" then income = 0
If assets = "" then assets = 0
If rent = "" then rent = 0

'Calculates expedited status based on above numbers
If (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) then expedited_status = "client does not appear expedited"
'----------------------------------------------------------------------------------------------------

'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
Call navigate_to_MAXIS_screen("STAT", "DISQ")
'grabbing footer month and year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Reads the DISQ info for the case note.
EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then 
	has_DISQ = False
Else
	has_DISQ = True
End if

'Reads MONY/DISB to see if EBT account is open 
IF expedited_status = "client appears expedited" THEN 
	Call navigate_to_MAXIS_screen("MONY", "DISB")
	EMReadScreen EBT_account_status, 1, 14, 27
END IF 

'THE CASE NOTE----------------------------------------------------------------------------------------------------
	call navigate_to_screen("case", "note")
	PF9
	
	EMReadScreen case_note_check, 17, 2, 33
	EMReadScreen mode_check, 1, 20, 09
	If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then    'this will account for those cases when the script is run on an out of county case.
		msgbox "The script can't open a case note. You may be in inquiry or entered a case number that is in another county." &_
		vbNewLine & vbNewLine & "This result for this case is " & expedited_status & vbNewLine & vbNewLine & "Please run the script again if you were in inquiry to add a case note."
		script_end_procedure("")
	else	
		'Body of the case note 
		Call write_variable_in_CASE_NOTE("Received " & application_type & ", " & expedited_status)
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
		call write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
		call write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
		call write_variable_in_CASE_NOTE("        Utilities (amt/HEST claimed): $" & utilities)
		call write_variable_in_CASE_NOTE("---")
		If has_DISQ = True then call write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
		If has_DISQ = False then call write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "Y" then call write_variable_in_CASE_NOTE("* EBT Account IS open.  Recipient will NOT be able to get a replacement card in the agency.  Rapid Electronic Issuance (REI) with caution.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "N" then call write_variable_in_CASE_NOTE("* EBT Account is NOT open.  Recipient is able to get initial card in the agency.  Rapid Electronic Issuance (REI) can be used, but only to avoid an emergency issuance or to meet EXP criteria.")
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)
		If expedited_status = "client appears expedited" then
			MsgBox "This client appears expedited. A same day interview needs to be offered."
		End if
		If expedited_status = "client does not appear expedited" then
			MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
		End if
	End if
script_end_procedure("")


