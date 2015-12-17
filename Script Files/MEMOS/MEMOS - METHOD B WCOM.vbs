'STATS GATHERING---------------------------------------------------------------------------------------------------- 
 name_of_script = "LTC method B SPEC - WCOM breakdown.vbs" 
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

BeginDialog MEMOS_LTC_METHOD_B_dialog, 0, 0, 271, 245, "LTC method B SPEC - WCOM breakdown"
  EditBox 85, 10, 75, 20, case_number
  EditBox 85, 35, 75, 20, Income
  EditBox 85, 60, 75, 20, MA_income_standard
  EditBox 85, 85, 75, 20, Medicare_B
  EditBox 85, 110, 75, 20, Medicare_D
  EditBox 85, 135, 75, 20, Dental_Premium
  EditBox 85, 160, 75, 20, HC_Premium
  EditBox 85, 185, 75, 20, Remedial_Care
  ButtonGroup ButtonPressed
    OkButton 60, 220, 50, 15
    CancelButton 120, 220, 50, 15
  Text 10, 185, 70, 15, "Remedial Care      Enter amount here"
  Text 5, 115, 70, 15, "Medicare Part D"
  Text 5, 40, 40, 10, "Income"
  Text 5, 65, 70, 15, " MA income standard "
  Text 5, 15, 70, 15, "Case Number"
  Text 5, 90, 70, 15, "Medicare Part B"
  Text 10, 135, 70, 15, "Dental Premium      Enter amount here"
  Text 10, 160, 70, 15, "Health Care Premium Enter amount here"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------


'Connects to BlueZone
EMConnect ""

'Finds the case number
row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then 
  EMReadScreen case_number, 8, row, col + 10
  case_number = replace(case_number, "_", "")
  case_number = trim(case_number)
End if


'GETTING DATA FROM MEDI

call navigate_to_screen("STAT", "MEDI")

EMReadScreen Medicare_B, 6, 7, 48
EMReadScreen Medicare_D, 6, 7, 75

IF Medicare_B = "______" then Medicare_B = "0"
IF Medicare_D = "______" then Medicare_D = "0"


'GETTING INCOME STANDARD AND SPENDOWN AMOUNTS

call navigate_to_screen("ELIG", "HC")
EMSendKey "x"
transmit

EMReadScreen MA_income_standard, 7, 16, 19

EMReadScreen Income, 7, 15, 19

IF Income = " " then Income = 0



'Shows the dialog

  Do
    Do
      Dialog MEMOS_LTC_METHOD_B_dialog
      If buttonpressed = 0 then stopscript
      If isnumeric(case_number) = False then MsgBox "You must enter a valid MAXIS case number."
    Loop until (isnumeric(case_number) = True) 
    transmit
    If isnumeric(case_number) = True then
      EMReadScreen MAXIS_check, 5, 1, 39
      If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your screen to MAXIS and try again. You might be passworded out."
    End if
  Loop until MAXIS_check = "MAXIS"



'THE MEMO

CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting HC notice

wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "HC" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			exit Do
		End If
	End If
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_hc_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_hc_waiting = true then script_end_procedure("No waiting HC results were found for the requested month")


Transmit
PF9

IF Dental_Premium = "" then Dental_Premium = 0.00
IF HC_Premium = "" then HC_Premium = 0.00
IF Remedial_Care = "" then Remedial_Care = 0.00
IF Medicare_B = "" then Medicare_B = 0.00
IF Medicare_D = "" then Medicare_D = 0.00


spenddown = Income - MA_income_standard
recipient_amount = spenddown - Medicare_B - Medicare_D - Dental_Premium - HC_Premium - Remedial_Care


'Worker Comment Input

EMSendKey "Although your spenddown is $" & spenddown & " your recipient amount the amount that you are responsible to pay each month is $" & recipient_amount
EMSendKey "<Tab>"
EMSendKey "This was determined using the following calculations:"
EMSendKey "<Tab>"
EMSendKey "<Tab>"
EMSendKey "Income: $" & Income & " - MA Income Standard $" & MA_income_standard & " = $" & spenddown & " Spenddown"

EMSendKey "<Tab>"
EMSendKey "<Tab>"
EMSendKey "Spenddown             $" & spenddown 
EMSendKey "<Tab>"
EMSendKey "Medicare B          - $" & Medicare_B
EMSendKey "<Tab>"
EMSendKey "Medicare D          - $" & Medicare_D
EMSendKey "<Tab>"
EMSendKey "Dental Premium      - $" & Dental_Premium
EMSendKey "<Tab>"
EMSendKey "Health Care Premium - $" & HC_Premium
EMSendKey "<Tab>"
EMSendKey "Remedial Care       - $" & Remedial_Care
EMSendKey "<Tab>"
EMSendKey "Recipient Amount    = $" & recipient_amount







