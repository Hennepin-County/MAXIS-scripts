'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - MA-EPD EI FIAT.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 600                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
current_month_plus_one = dateadd("m", 1, date)

MAXIS_footer_month = datepart("m", current_month_plus_one)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month

MAXIS_footer_year = datepart("yyyy", current_month_plus_one)
MAXIS_footer_year = MAXIS_footer_year - 2000

current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month

current_year = datepart("yyyy", date)
current_year = current_year - 2000

current_month_and_year = current_month & "/" & current_year
next_month_and_year = MAXIS_footer_month & "/" & MAXIS_footer_year

'DIALOGS--------------------------------
BeginDialog case_number_dialog, 0, 0, 156, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 90, 0, 60, 15, MAXIS_case_number
  Text 25, 25, 65, 10, "HH memb number:"
  EditBox 90, 20, 30, 15, memb_number
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

'THE SCRIPT

EMConnect ""

call MAXIS_case_number_finder(MAXIS_case_number)

memb_number = "01" 'Setting a default

Dialog case_number_dialog
Cancel_confirmation

call navigate_to_MAXIS_screen("STAT", "JOBS")
EMReadScreen jobs_memb, 2, 4, 33  'checking if the current jobs panel is the memb, if not it will nav to member.
IF jobs_memb <> memb_number THEN
	EMWriteScreen memb_number, 20, 76
	transmit
END IF
EMReadScreen jobs_total, 1, 2, 78
EMReadScreen jobs_current, 1, 2, 73

If jobs_total = "0" then MsgBox "No JOBS panel is known for this client. You will have to enter income amounts manually."

If jobs_current = "1" then
  EMReadScreen pay_freq_01, 1, 18, 35
  If pay_freq_01 = "1" then frequency_job_01 = "1: monthly"
  If pay_freq_01 = "2" then frequency_job_01 = "2: twice monthly"
  If pay_freq_01 = "3" then frequency_job_01 = "3: every 2 weeks"
  If pay_freq_01 = "4" then frequency_job_01 = "4. every week"
  If pay_freq_01 = "5" then frequency_job_01 = "5. other (use monthly avg)"
  EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  IF HC_income_est_check = "Est" Then 'this is the old position
	EMWriteScreen "x", 19, 54
  ELSE								'this is the new position
	EMWriteScreen "x", 19, 48
  END IF
  transmit
  EMReadScreen income_job_01, 8, 11, 63
  income_job_01 = trim(replace(income_job_01, "_", ""))
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "2" then
  EMReadScreen pay_freq_02, 1, 18, 35
  If pay_freq_02 = "1" then frequency_job_02 = "1: monthly"
  If pay_freq_02 = "2" then frequency_job_02 = "2: twice monthly"
  If pay_freq_02 = "3" then frequency_job_02 = "3: every 2 weeks"
  If pay_freq_02 = "4" then frequency_job_02 = "4. every week"
  If pay_freq_02 = "5" then frequency_job_02 = "5. other (use monthly avg)"
  EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  IF HC_income_est_check = "Est" Then 'this is the old position
	EMWriteScreen "x", 19, 54
  ELSE								'this is the new position
	EMWriteScreen "x", 19, 48
  END IF
  transmit
  EMReadScreen income_job_02, 8, 11, 63
  income_job_02 = trim(replace(income_job_02, "_", ""))
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If jobs_current = "3" then
  EMReadScreen pay_freq_03, 1, 18, 35
  If pay_freq_03 = "1" then frequency_job_03 = "1: monthly"
  If pay_freq_03 = "2" then frequency_job_03 = "2: twice monthly"
  If pay_freq_03 = "3" then frequency_job_03 = "3: every 2 weeks"
  If pay_freq_03 = "4" then frequency_job_03 = "4. every week"
  If pay_freq_03 = "5" then frequency_job_03 = "5. other (use monthly avg)"
  EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
  IF HC_income_est_check = "Est" Then 'this is the old position
	EMWriteScreen "x", 19, 54
  ELSE								'this is the new position
	EMWriteScreen "x", 19, 48
  END IF
  transmit
  EMReadScreen income_job_03, 8, 11, 63
  income_job_03 = trim(replace(income_job_03, "_", ""))
  transmit
  transmit
  EMReadScreen jobs_current, 1, 2, 73
End if

If income_job_01 = "" then
  income_job_01 = income_job_02
  frequency_job_01 = frequency_job_02
  income_job_02 = ""
  frequency_job_02 = ""
End if

If income_job_02 = "" then
  income_job_02 = income_job_03
  frequency_job_02 = frequency_job_03
  income_job_03 = ""
  frequency_job_03 = ""
End if

BeginDialog MA_EPD_dialog, 0, 0, 186, 156, "MA-EPD dialog"
  Text 35, 5, 40, 10, "Income amt"
  Text 115, 5, 30, 10, "Pay freq."
  Text 5, 25, 25, 10, "Job 1:"
  EditBox 30, 20, 40, 15, income_job_01
  DropListBox 85, 20, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_01
  Text 5, 45, 25, 10, "Job 2:"
  EditBox 30, 40, 40, 15, income_job_02
  DropListBox 85, 40, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_02
  Text 5, 65, 25, 10, "Job 3:"
  EditBox 30, 60, 40, 15, income_job_03
  DropListBox 85, 60, 90, 15, " "+chr(9)+"1: monthly"+chr(9)+"2: twice monthly"+chr(9)+"3: every 2 weeks"+chr(9)+"4. every week"+chr(9)+"5. other (use monthly avg)", frequency_job_03
  GroupBox 20, 85, 140, 40, "Script should update:"
  OptionGroup RadioGroup1
    RadioButton 25, 95, 110, 10, "Current and future months", Radio1
    RadioButton 25, 110, 100, 10, "Just future months", Radio2
  ButtonGroup ButtonPressed
    OkButton 40, 135, 50, 15
    CancelButton 100, 135, 50, 15
EndDialog

Dialog MA_EPD_dialog
cancel_confirmation

'SECTION 04: NOW IT GOES TO ELIG/HC TO FIAT THE AMOUNTS
Call navigate_to_MAXIS_screen("ELIG", "HC__")

row = 1
col = 1
EMSearch memb_number & " ", row, col 'finding the member number
If row = 0 then
  MsgBox "Member number not found. You may have entered an incorrect member number on the first screen. Try the script again."
  StopScript
End if

EMWriteScreen "x", row, 26
transmit

EMReadScreen elig_type_check_first_month, 2, 12, 17
EMReadScreen elig_type_check_second_month, 2, 12, 28
EMReadScreen elig_type_check_third_month, 2, 12, 39
EMReadScreen elig_type_check_fourth_month, 2, 12, 50
EMReadScreen elig_type_check_fifth_month, 2, 12, 61
EMReadScreen elig_type_check_sixth_month, 2, 12, 72

If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then MsgBox "Not all of the months of this case are MA-EPD. Process manually."
If elig_type_check_first_month <> "DP" and elig_type_check_second_month <> "DP" and elig_type_check_third_month <> "DP" and elig_type_check_fourth_month <> "DP" and elig_type_check_fifth_month <> "DP" and elig_type_check_sixth_month <> "DP" then stopscript

PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "05"
  transmit
End if
If radio1 = 1 then
  row = 6
  col = 1
  EMSearch current_month_and_year, row, col
End if

If radio2 = 1 or row = 0 then
  row = 6
  col = 1
  EMSearch next_month_and_year, row, col
End if

'Multiplier calculations
If frequency_job_01 = "1: monthly" or frequency_job_01 = "5. other (use monthly avg)" then multiplier_01 = 1
If frequency_job_02 = "1: monthly" or frequency_job_02 = "5. other (use monthly avg)" then multiplier_02 = 1
If frequency_job_03 = "1: monthly" or frequency_job_03 = "5. other (use monthly avg)" then multiplier_03 = 1

If frequency_job_01 = "2: twice monthly" then multiplier_01 = 2
If frequency_job_02 = "2: twice monthly" then multiplier_02 = 2
If frequency_job_03 = "2: twice monthly" then multiplier_03 = 2

If frequency_job_01 = "3: every 2 weeks" then multiplier_01 = 2.16
If frequency_job_02 = "3: every 2 weeks" then multiplier_02 = 2.16
If frequency_job_03 = "3: every 2 weeks" then multiplier_03 = 2.16

If frequency_job_01 = "4. every week" then multiplier_01 = 4.3
If frequency_job_02 = "4. every week" then multiplier_02 = 4.3
If frequency_job_03 = "4. every week" then multiplier_03 = 4.3

Do
  EMWriteScreen "x", 9, col + 2
  transmit
  EMWriteScreen "x", 13, 03
  transmit
  EMWriteScreen "___________", 8, 43
  EMWriteScreen income_job_01 * multiplier_01, 8, 43
  If income_job_02 <> "" then
    EMWriteScreen "___________", 9, 43
    EMWriteScreen income_job_02 * multiplier_02, 9, 43
  End if
  If income_job_03 <> "" then
    EMWriteScreen "___________", 10, 43
    EMWriteScreen income_job_03 * multiplier_03, 10, 43
  End if
  col = col + 11
  transmit
  transmit
  transmit
loop until col > 76

script_end_procedure("Success! Please make sure to check eligibility for any Medicare savings programs such as QMB or SLMB.")
