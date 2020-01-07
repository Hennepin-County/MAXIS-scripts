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
call changelog_update("03/28/2018", "Updated dialog box and functionality.", "MiKayla Handley")
call changelog_update("09/25/2017", "Updated income standards for 130% FPG effective 10/17. Also updated error message handling on the back end.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, and finding case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'T'ODO Multiple-Case noted
'Dialog is presented. Requires all sections other than spousal sponsor income to be filled out.
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 346, 290, "Sponsor income calculation"
  EditBox 55, 5, 40, 15, MAXIS_case_number
  EditBox 160, 5, 20, 15, memb_number
  CheckBox 190, 5, 105, 10, "Verified via SAVE requested?", via_save
  CheckBox 190, 15, 135, 10, "TPQY / SSA quarters checked? HRS:", TPQY_check
  EditBox 320, 10, 20, 15, SSA_hours
  EditBox 45, 35, 55, 15, primary_sponsor_earned_income
  EditBox 145, 35, 55, 15, spousal_sponsor_earned_income
  DropListBox 260, 35, 75, 15, "Select One:"+chr(9)+"Paystubs"+chr(9)+"Taxes"+chr(9)+"EVF"+chr(9)+"SMI"+chr(9)+"Other please specify ", earned_income_verification
  EditBox 45, 70, 55, 15, primary_sponsor_unearned_income
  EditBox 145, 70, 55, 15, spousal_sponsor_unearned_income
  DropListBox 260, 70, 75, 15, "Select One:"+chr(9)+"Paystubs"+chr(9)+"Taxes"+chr(9)+"EVF"+chr(9)+"SMI"+chr(9)+"Other please specify ", unearned_income_verification
  EditBox 80, 105, 70, 15, name_sponsor
  EditBox 255, 105, 80, 15, name_of_spon_spouse
  EditBox 80, 125, 165, 15, sponsor_addr
  EditBox 280, 125, 55, 15, phone_one
  EditBox 80, 145, 20, 15, spon_HH_size
  EditBox 205, 145, 20, 15, number_of_spon_clients
  EditBox 80, 170, 70, 15, name_sponsor_two
  EditBox 255, 170, 80, 15, name_of_spon_spouse_two
  EditBox 80, 190, 165, 15, sponsor_addr_two
  EditBox 280, 190, 55, 15, phone_two
  EditBox 80, 210, 20, 15, spon_HH_size_two
  EditBox 205, 210, 20, 15, number_of_spon_clients_two
  EditBox 200, 245, 135, 15, denial_reason
  EditBox 60, 270, 175, 15, other_notes
  CheckBox 10, 240, 110, 10, "Indigent Exemption Reviewed?", indexmp_CHECKBOX
  CheckBox 10, 250, 85, 10, "DV Waiver Reviewed?", DVW_CHECKBOX
  CheckBox 245, 215, 90, 10, "Check if additional SPON", additonal_spon_CHECKBOX
  ButtonGroup ButtonPressed
    OkButton 245, 270, 45, 15
    CancelButton 295, 270, 45, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 100, 10, 60, 10, "Member Number:"
  GroupBox 5, 25, 335, 30, "Earned income to deem:"
  Text 15, 40, 30, 10, "Primary:"
  Text 110, 40, 30, 10, "Spousal:"
  Text 205, 40, 55, 10, "Income verified: "
  GroupBox 5, 60, 335, 30, "Unearned income to deem:"
  Text 15, 75, 30, 10, "Primary:"
  Text 110, 75, 30, 10, "Spousal:"
  Text 205, 75, 55, 10, "Income verified: "
  GroupBox 5, 95, 335, 135, "Sponsor Information:"
  Text 15, 110, 60, 10, "Name of sponsor:"
  Text 160, 110, 90, 10, "Name of sponsor's spouse:"
  Text 45, 130, 30, 10, "Address:"
  Text 250, 130, 25, 10, "Phone:"
  Text 15, 150, 60, 10, "Sponsor HH size:"
  Text 105, 150, 100, 10, "Number of sponsored clients:"
  Text 15, 175, 60, 10, "Name of sponsor:"
  Text 160, 175, 90, 10, "Name of sponsor's spouse:"
  Text 45, 195, 30, 10, "Address:"
  Text 255, 195, 25, 10, "Phone:"
  Text 15, 215, 60, 10, "Sponsor HH size:"
  Text 105, 215, 100, 10, "Number of sponsored clients:"
  GroupBox 5, 230, 335, 35, "HP. Immigration Information"
  Text 135, 250, 65, 10, "Reason for denial?"
  Text 15, 275, 45, 10, "Other Notes:"
ENDDIALOG
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbCr & "* You must enter a valid case number."
		If isnumeric(primary_sponsor_earned_income) = False and isnumeric(spousal_sponsor_earned_income) = False and isnumeric(primary_sponsor_unearned_income) = False and isnumeric(spousal_sponsor_unearned_income) = False THEN err_msg = err_msg & vbCr & "* You must enter some income. You can enter a ''0'' if that is accurate."
		If isnumeric(sponsor_HH_size) = False THEN err_msg = err_msg & vbCr & "* You must enter a sponsor HH size."
		If isnumeric(number_of_spon_clients) = False THEN err_msg = err_msg & vbCr & "* You must enter the number of sponsored clients."
		If worker_signature = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
'Determines the income limits
' >> Income limits from CM 19.06 - MAXIS Gross Income 130% FPG (Updated effective 10/01/18)
If date >= cdate("10/01/2018") then
    If sponsor_HH_size = 1 then income_limit = 1316
    If sponsor_HH_size = 2 then income_limit = 1784
    If sponsor_HH_size = 3 then income_limit = 2252
    If sponsor_HH_size = 4 then income_limit = 2720
    If sponsor_HH_size = 5 then income_limit = 3188
    If sponsor_HH_size = 6 then income_limit = 3656
    If sponsor_HH_size = 7 then income_limit = 4124
    If sponsor_HH_size = 8 then income_limit = 4592
    If sponsor_HH_size > 8 then income_limit = 4592 + (468 * (sponsor_HH_size - 8))
else
    If sponsor_HH_size = 1 then income_limit = 1307
    If sponsor_HH_size = 2 then income_limit = 1760
    If sponsor_HH_size = 3 then income_limit = 2213
    If sponsor_HH_size = 4 then income_limit = 2665
    If sponsor_HH_size = 5 then income_limit = 3118
    If sponsor_HH_size = 6 then income_limit = 3571
    If sponsor_HH_size = 7 then income_limit = 4024
    If sponsor_HH_size = 8 then income_limit = 4477
    If sponsor_HH_size > 8 then income_limit = 4477 + (453 * (sponsor_HH_size - 8))
End if

'If any income variables are not numeric, the script will convert them to a "0" for calculating
If IsNumeric(primary_sponsor_earned_income) = False then primary_sponsor_earned_income = 0
If IsNumeric(spousal_sponsor_earned_income) = False then spousal_sponsor_earned_income = 0
If IsNumeric(primary_sponsor_unearned_income) = False then primary_sponsor_unearned_income = 0
If IsNumeric(spousal_sponsor_unearned_income) = False then spousal_sponsor_unearned_income = 0

'Determines the sponsor deeming amount for SNAP
SNAP_EI_disregard = (abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income)) * 0.2
sponsor_deeming_amount_SNAP = ((((abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income)) - SNAP_EI_disregard) + (abs(primary_sponsor_unearned_income) + abs(spousal_sponsor_unearned_income)) - income_limit)/abs(number_of_spon_clients))

'Determines the sponsor deeming amount for other programs
sponsor_deeming_amount_other_programs = abs(primary_sponsor_earned_income) + abs(spousal_sponsor_earned_income) + abs(primary_sponsor_unearned_income) + abs(spousal_sponsor_unearned_income)

'If the deeming amounts are less than 0 they need to show a 0
If sponsor_deeming_amount_SNAP < 0 then sponsor_deeming_amount_SNAP = 0
If sponsor_deeming_amount_other_programs < 0 then sponsor_deeming_amount_other_programs = 0
'phone_one, & "-" & phone_two & "-" & phone_three
'Case note the findings
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("SPON-Income deeming calculation for M" & memb_number)
If primary_sponsor_earned_income <> 0 then
	call write_bullet_and_variable_in_case_note("Primary sponsor earned income", "$" & primary_sponsor_earned_income)
	CALL write_bullet_and_variable_in_case_note("Income verification received", earned_income_verification)
END IF
If spousal_sponsor_earned_income <> 0 then call write_bullet_and_variable_in_case_note("Spousal sponsor earned income", "$" & spousal_sponsor_earned_income)
If primary_sponsor_unearned_income <> 0 then
	call write_bullet_and_variable_in_case_note("Primary sponsor unearned income", "$" & primary_sponsor_unearned_income)
	CALL write_bullet_and_variable_in_case_note("Unearned income verification received", unearned_income_verification)
END IF
If spousal_sponsor_unearned_income <> 0 then call write_bullet_and_variable_in_case_note("Spousal sponsor unearned income", "$" & spousal_sponsor_unearned_income)
If SNAP_EI_disregard <> 0 then call write_bullet_and_variable_in_case_note("20% diregard of EI for SNAP", "$" & SNAP_EI_disregard)
CALL write_bullet_and_variable_in_case_note("Sponsor HH size and income limit", sponsor_HH_size & ", $" & income_limit)
CALL write_bullet_and_variable_in_case_note("Number of sponsored people", number_of_spon_clients)
call write_bullet_and_variable_in_case_note("Sponsor deeming amount for SNAP", "$" & sponsor_deeming_amount_SNAP)
call write_bullet_and_variable_in_case_note("Sponsor deeming amount for other programs", "$" & sponsor_deeming_amount_other_programs)
CALL write_bullet_and_variable_in_case_note("Verified via SAVE requested?", via_save)
IF TPQY_check = CHECKED THEN CALL write_variable_in_case_note("TPQY/SSA quarters checked? HRS:")
CALL write_bullet_and_variable_in_case_note("TPQY/SSA quarters HRS", SSA_hours)
CALL write_bullet_and_variable_in_case_note("Name of Sponsor", name_sponsor)
CALL write_bullet_and_variable_in_case_note("Name of Sponsor's Spouse", name_of_spon_spouse)
CALL write_bullet_and_variable_in_case_note("Address", sponsor_addr)
CALL write_bullet_and_variable_in_case_note("Phone", phone_one)
CALL write_bullet_and_variable_in_case_note("Name of second sponsor", name_of_spon_spouse_two)
CALL write_bullet_and_variable_in_case_note("Address", sponsor_addr_two)
CALL write_bullet_and_variable_in_case_note("Phone", phone_two)
CALL write_bullet_and_variable_in_case_note("Second Sponsor HH size", spon_HH_size_two)
CALL write_bullet_and_variable_in_case_note("Reason for Denial", denial_reason)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF indexmp_CHECKBOX = CHECKED THEN CALL write_variable_in_case_note("Indigent Exemption Reviewed")
IF DVW_CHECKBOX = CHECKED THEN CALL write_bullet_and_variable_in_case_note("DV Waiver Reviewed?")
CALL write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Updated SPON income and casenote")
