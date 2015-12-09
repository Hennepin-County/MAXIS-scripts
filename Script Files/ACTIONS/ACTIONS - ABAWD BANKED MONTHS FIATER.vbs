
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - ABAWD BANKED MONTHS FIATER.vbs"
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

'Defining variables
'Dim gross_wages, busi_income, gross_RSDI, gross_SSI, gross_VA, gross_UC, gross_CS, gross_other
'Dim deduction_FMED, deduction_DCEX, deduction_COEX

'Dialog
BeginDialog Dialog1, 0, 0, 191, 130, "ABAWD BANKED MONTHS FIATER"
  ButtonGroup ButtonPressed
    OkButton 80, 110, 50, 15
    CancelButton 135, 110, 50, 15
  EditBox 70, 10, 80, 15, case_number
  EditBox 105, 30, 15, 15, initial_month
  EditBox 135, 30, 15, 15, initial_year
  Text 20, 15, 50, 10, "Case Number:"
  Text 20, 30, 75, 20, "Initial month / year of package:"
  Text 125, 35, 5, 10, "/"
  Text 10, 55, 175, 50, "This script will FIAT SNAP results for each month between initial month and current month plus one to make an approval for ABAWD Banked Months.  All STAT panels must be updated correctly before using this script to create FIAT results.  Refer to SNAP bulletin: "
EndDialog

'The script.

call 


'Create hh_member_array
call HH_member_custom_dialog(HH_member_array)
'for each member in hh_member_array
	'go to STAT / MEMB and pull member age
'Collecting the 
'defining necessary dates
initial_date = initial_month & "/01/" & initial_year
current_month = initial_date
current_month_plus_one = dateadd("m", date, 1) 
maxis_background_check
'The following loop will take the script throught each month in the package, from appl month. to CM+1
Do
	footer_month = datepart("m", current_month)
	if len(footer_month) = 1 THEN footer_month = "0" & footer_month 
	footer_year = right(datepart("YYYY", current_month), 2)
	
	'background check
	'for each member in hh_member_array
		'go to UNEA and read SNAP PIC for each thing
		'if member > 18 go to JOBS and read SNAP PIC
		'if member > 18 go to BUSI and read SNAP PIC
		'if member > 18 go to RBIC and read SNAP PIC
		'go to COEX and read deductions
		'go to DCEX and read deductions
		
	'Sum up gross income
	'background check
	maxis_background_check
	'Go to FIAT
	back_to_self
	EMwritescreen "FIAT", 16, 43
	EMWritescreen case_number, 18, 43
	EMwritescreen footer_month, 20, 43
	EMWritescreen footer_year, 20, 46 
	msgbox "WTF?"
	transmit
	EMReadscreen results_check, 4, 14, 46 'We need to make sure results exist, otherwise stop.
	IF results_check = "    " THEN script_end_procedure("The script was unable to find unapproved SNAP results for the benefit month, please check your case and try again.")
	EMWritescreen "03", 4, 34 'entering the FIAT reason
	EMWritescreen "x", 14, 22
	transmit 'This should take us to FFSL
	'The following loop will enter person tests screen and pass for each member on grant
	abawd_check = false
	For each member in hh_member_array
		row = 6
		col = 1
		EMSearch member, row, col 'Finding the row this member is on
		EMWritescreen "x", row, 5
		transmit 'Now on FFPR
		EMReadscreen inelig_test, 6, 9, 12 'This is error proofing to make sure the case has at least 1 ineligible member
		IF inelig_test = "FAILED" THEN abawd_check = True
		EMWritescreen "PASSED", 9, 12
		transmit
		PF3 'back to FFSL
	Next
	IF abawd_check = false THEN script_end_procedure("ERROR: There are no members on this case with ineligible ABAWDs.  The script will stop.")
	'Ready to head into case test / budget screens
	EMWritescreen "x", 16, 5
	EMWritescreen "x", 17, 5
	Transmit
	'Passing all case tests  
			'Need to add logic to determine which tests to pass.
	EMWritescreen "PASSED", 10, 7
	EMWritescreen "PASSED", 13, 7
	EMWritescreen "PASSED", 14, 7
	PF3
	'Now the BUDGET (FFB1) NO 
	EMWritescreen gross_wages, 5, 32
	EMWritescreen busi_income, 6, 32
	EMWritescreen gross_RSDI, 11, 32
	EMWritescreen gross_SSI, 12, 32
	EMWritescreen gross_VA, 13, 32
	EMWritescreen gross_UC, 14, 32
	EMWritescreen gross_CS, 15, 32
	EMWritescreen gross_other, 16, 32
	EMWritescreen deduction_FMED, 12, 72
	EMWritescreen deduction_DCEX, 13, 72
	EMWritescreen deduction_COEX, 14, 72
	transmit
	'Now on FFB2
	EMWritescreen SHEL_rent, 5, 29
	EMWritescreen SHEL_tax, 6, 29
	EMWritescreen SHEL_insa, 7, 29
	EMWritescreen HEST_elect, 8, 29
	EMWritescreen HEST_heat, 9, 29
	EMWritescreen HEST_phone, 10, 29
	transmit
	'Now on SUMM screen, which shouldn't matter
	PF3 'back to FFSL
	PF3 'This should bring up the "do you want to retain" popup
	'should add error proofing here to check for the "summ needs to be last action" or whatever it is popup
	EMWritescreen "Y", 13, 41
	transmit
	EMReadscreen final_month_check, 10, 53 'This looks for a popup that only comes up in the final month, and clears it.
	IF final_month_check = "ELIG" THEN
		EMWritescreen "Y", 11, 52
		EMWritescreen month_array(0), 13, 37
		transmit
	END IF
	IF current_month = current_month_plus_one THEN exit DO
	current_month = dateadd("m", current_month, 1) 
Loop  

script_end_procedure("Success. The FIAT results have been generated. Please review before approving.")
