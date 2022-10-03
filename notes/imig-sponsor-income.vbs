'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - IMIG - SPONSOR INCOME.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("10/03/2022", "Updated income standards for 130% FPG effective 10/22.", "Ilse Ferris, Hennepin County")
call changelog_update("09/29/2021", "Updated income standards for 130% FPG effective 10/21.", "Ilse Ferris, Hennepin County")
call changelog_update("10/01/2020", "Updated income standards for 130% FPG effective 10/20.", "Ilse Ferris, Hennepin County")
call changelog_update("10/01/2019", "Updated income standards for 130% FPG effective 10/19.", "Ilse Ferris, Hennepin County")
call changelog_update("09/01/2018", "Updated income standards for 130% FPG effective 10/18.", "Ilse Ferris, Hennepin County")
call changelog_update("09/25/2017", "Updated income standards for 130% FPG effective 10/17. Also updated error message handling on the back end.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 216, 165, "Sponsor income calculation dialog"
  EditBox 65, 10, 70, 15, MAXIS_case_number
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
'Connecting to BlueZone, and finding case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Dialog is presented. Requires all sections other than spousal sponsor income to be filled out.
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbCr & "* You must enter a valid case number."
		If isnumeric(primary_sponsor_earned_income) = False and isnumeric(spousal_sponsor_earned_income) = False and isnumeric(primary_sponsor_unearned_income) = False and isnumeric(spousal_sponsor_unearned_income) = False THEN err_msg = err_msg & vbCr & "* You must enter some income. You can enter a ''0'' if that is accurate."
		If isnumeric(sponsor_HH_size) = False THEN err_msg = err_msg & vbCr & "* You must enter a sponsor HH size."
		If isnumeric(number_of_sponsored_immigrants) = False THEN err_msg = err_msg & vbCr & "* You must enter the number of sponsored immigrants."
		If worker_signature = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Determines the income limits
' >> Income limits from CM 19.06 - MAXIS Gross Income 130% FPG at: https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_001906
If DateDiff("d",date,#10/01/2022#) <= 0 then
    'October 2022 -- Amounts for applications on or AFTER 10/01/2022
    If sponsor_HH_size = 1 then income_limit = 1473
    If sponsor_HH_size = 2 then income_limit = 1984
    If sponsor_HH_size = 3 then income_limit = 2495
    If sponsor_HH_size = 4 then income_limit = 3007
    If sponsor_HH_size = 5 then income_limit = 3518
    If sponsor_HH_size = 6 then income_limit = 4029
    If sponsor_HH_size = 7 then income_limit = 4541
    If sponsor_HH_size = 8 then income_limit = 5052
    If sponsor_HH_size > 8 then income_limit = 5052 + (512 * (sponsor_HH_size - 8))
Elseif DateDiff("d",date,#10/01/2022#) > 0 then
    'October 2021 -- Amounts for applications on or BEFORE 10/01/2022
    If sponsor_HH_size = 1 then income_limit = 1396
    If sponsor_HH_size = 2 then income_limit = 1888
    If sponsor_HH_size = 3 then income_limit = 2379
    If sponsor_HH_size = 4 then income_limit = 2871
    If sponsor_HH_size = 5 then income_limit = 3363
    If sponsor_HH_size = 6 then income_limit = 3855
    If sponsor_HH_size = 7 then income_limit = 4347
    If sponsor_HH_size = 8 then income_limit = 4839
    If sponsor_HH_size > 8 then income_limit = 4839 + (492 * (sponsor_HH_size - 8))
End if

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
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("~~~Sponsor deeming income calculation~~~")
If primary_sponsor_earned_income <> 0 then call write_bullet_and_variable_in_case_note("Primary sponsor earned income", "$" & primary_sponsor_earned_income)
If spousal_sponsor_earned_income <> 0 then call write_bullet_and_variable_in_case_note("Spousal sponsor earned income", "$" & spousal_sponsor_earned_income)
If primary_sponsor_unearned_income <> 0 then call write_bullet_and_variable_in_case_note("Primary sponsor unearned income", "$" & primary_sponsor_unearned_income)
If spousal_sponsor_unearned_income <> 0 then call write_bullet_and_variable_in_case_note("Spousal sponsor unearned income", "$" & spousal_sponsor_unearned_income)
If SNAP_EI_disregard <> 0 then call write_bullet_and_variable_in_case_note("20% diregard of EI for SNAP", "$" & SNAP_EI_disregard)
call write_bullet_and_variable_in_case_note("Sponsor HH size and income limit", sponsor_HH_size & ", $" & income_limit)
call write_bullet_and_variable_in_case_note("Number of sponsored immigrants", number_of_sponsored_immigrants)
call write_bullet_and_variable_in_case_note("Sponsor deeming amount for SNAP", "$" & sponsor_deeming_amount_SNAP)
call write_bullet_and_variable_in_case_note("Sponsor deeming amount for other programs", "$" & sponsor_deeming_amount_other_programs)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
