'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LEP - SPONSOR INCOME.vbs"
start_time = timer


'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF


'DIALOGS--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog sponsor_income_calculation_dialog, 0, 0, 216, 165, "Sponsor income calculation dialog"
  EditBox 65, 10, 70, 15, case_number
  EditBox 40, 45, 55, 15, primary_sponsor_earned_income
  EditBox 150, 45, 55, 15, spousal_sponsor_earned_income
  EditBox 40, 80, 55, 15, primary_sponsor_unearned_income
  EditBox 150, 80, 55, 15, spousal_sponsor_unearned_income
  EditBox 70, 105, 30, 15, sponsor_HH_size
  EditBox 120, 125, 30, 15, number_of_sponsored_immigrants
  EditBox 70, 145, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 125, 50, 15
    CancelButton 160, 145, 50, 15
  Text 10, 15, 50, 10, "Case number:"
  GroupBox 5, 35, 205, 30, "Earned income to deem:"
  Text 10, 50, 30, 10, "Primary:"
  Text 120, 50, 30, 10, "Spousal:"
  GroupBox 5, 70, 205, 30, "Unearned income to deem:"
  Text 10, 85, 30, 10, "Primary:"
  Text 120, 85, 30, 10, "Spousal:"
  Text 5, 110, 60, 10, "Sponsor HH size:"
  Text 5, 130, 115, 10, "Number of sponsored immigrants:"
  Text 5, 150, 65, 10, "Worker signature:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""

'Dialog is presented. Requires all sections other than spousal sponsor income to be filled out.
Do
  Do
    Do
      Do
        Do
          Do
            Dialog sponsor_income_calculation_dialog
            If ButtonPressed = 0 then stopscript
            If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must enter a valid case number."
          Loop until isnumeric(case_number) = True and len(case_number) <= 8
          If isnumeric(primary_sponsor_earned_income) = False and isnumeric(spousal_sponsor_earned_income) = False and isnumeric(primary_sponsor_unearned_income) = False and isnumeric(spousal_sponsor_unearned_income) = False then MsgBox "You must enter some income. You can enter a ''0'' if that is accurate."
        Loop until isnumeric(primary_sponsor_earned_income) = True or isnumeric(spousal_sponsor_earned_income) = True or isnumeric(primary_sponsor_unearned_income) = True or isnumeric(spousal_sponsor_unearned_income) = True
        If isnumeric(sponsor_HH_size) = False then MsgBox "You must enter a sponsor HH size."
      Loop until isnumeric(sponsor_HH_size) = True
      If isnumeric(number_of_sponsored_immigrants) = False then MsgBox "You must enter the number of sponsored immigrants."
    Loop until isnumeric(number_of_sponsored_immigrants) = True
    If worker_signature = "" then MsgBox "You must sign your case note!"
  Loop until worker_signature <> ""
  transmit
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "MAXIS not found. You might be locked out of your case. Check BlueZone and try again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Determines the income limits
If sponsor_HH_size = 1 then income_limit = 1265
If sponsor_HH_size = 2 then income_limit = 1705
If sponsor_HH_size = 3 then income_limit = 2144
If sponsor_HH_size = 4 then income_limit = 2584
If sponsor_HH_size = 5 then income_limit = 3024
If sponsor_HH_size = 6 then income_limit = 3464
If sponsor_HH_size = 7 then income_limit = 3904
If sponsor_HH_size = 8 then income_limit = 4344
If sponsor_HH_size > 8 then income_limit = 4344 + (440 * (sponsor_HH_size - 8))

'If any income variables are not numeric, the script will convert them to a "0" for calculating
If IsNumeric(primary_sponsor_earned_income) = False then primary_sponsor_earned_income = 0
If IsNumeric(spousal_sponsor_earned_income) = False then spousal_sponsor_earned_income = 0
If IsNumeric(primary_sponsor_unearned_income) = False then primary_sponsor_unearned_income = 0
If IsNumeric(spousal_sponsor_unearned_income) = False then spousal_sponsor_unearned_income = 0

'Determines the sponsor deeming amount for SNAP
SNAP_EI_disregard = (abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income)) * 0.2
sponsor_deeming_amount_SNAP = ((((abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income)) - SNAP_EI_disregard) + (abs(primary_sponsor_unearned_income) + abs(spousal_sponsor_unearned_income)) - income_limit)/abs(number_of_sponsored_immigrants))

'Determines the sponsor deeming amount for other programs
sponsor_deeming_amount_other_programs = abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income) + abs(primary_sponsor_unearned_income) + abs(spousal_sponsor_unearned_income)

'If the deeming amounts are less than 0 they need to show a 0
If sponsor_deeming_amount_SNAP < 0 then sponsor_deeming_amount_SNAP = 0
If sponsor_deeming_amount_other_programs < 0 then sponsor_deeming_amount_other_programs = 0

'Case note the findings
call navigate_to_screen("case", "note")
PF9
EMSendKey "~~~Sponsor deeming income calculation~~~" & "<newline>"
If primary_sponsor_earned_income <> 0 then call write_editbox_in_case_note("Primary sponsor earned income", "$" & primary_sponsor_earned_income, 6)
If spousal_sponsor_earned_income <> 0 then call write_editbox_in_case_note("Spousal sponsor earned income", "$" & spousal_sponsor_earned_income, 6)
If primary_sponsor_unearned_income <> 0 then call write_editbox_in_case_note("Primary sponsor unearned income", "$" & primary_sponsor_unearned_income, 6)
If spousal_sponsor_unearned_income <> 0 then call write_editbox_in_case_note("Spousal sponsor unearned income", "$" & spousal_sponsor_unearned_income, 6)
If SNAP_EI_disregard <> 0 then call write_editbox_in_case_note("20% diregard of EI for SNAP", "$" & SNAP_EI_disregard, 6)
call write_editbox_in_case_note("Sponsor HH size and income limit", sponsor_HH_size & ", $" & income_limit, 6)
call write_editbox_in_case_note("Number of sponsored immigrants", number_of_sponsored_immigrants, 6)
call write_editbox_in_case_note("Sponsor deeming amount for SNAP", "$" & sponsor_deeming_amount_SNAP, 6)
call write_editbox_in_case_note("Sponsor deeming amount for other programs", "$" & sponsor_deeming_amount_other_programs, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")






