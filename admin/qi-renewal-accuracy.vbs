'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - QI RENEWAL ACCURACY.vbs"
start_time = timer
STATS_counter = 1                   'sets the stats counter at one
STATS_manualtime = 70              'manual run time in seconds
STATS_denomination = "I"       		'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("07/27/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------FUNCTIONS
FUNCTION build_hh_array(hh_array)
	hh_array = ""
	panel_row = 5
	DO
		EMReadScreen person, 2, panel_row, 3
		IF trim(person) <> "" THEN
			hh_array = hh_array & person & ","
			panel_row = panel_row + 1
		END IF
	LOOP UNTIL trim(person) = ""
	hh_array = trim(hh_array)
	If right(hh_array, 1) = "," then hh_array = left(hh_array, len(hh_array) - 1)
	hh_array = split(hh_array, ",")
END FUNCTION

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function

Function updated_panel_member_array(stat_panel, output_variable)
    CALL navigate_to_MAXIS_screen("STAT", stat_panel)
    CALL build_hh_array(JOBS_array)
    FOR EACH HH_member IN JOBS_array
    	'MsgBox "HH memb: " & HH_member
    	IF HH_member <> "" THEN
    		CALL write_value_and_transmit(HH_member, 20, 76)
    		EMReadScreen updated_date, 8, 21, 55
    		updated_date = replace(updated_date, " ", "/")
			STATS_COUNTER = STATS_counter + 1                      'adds one instance to the stats counter
    		IF updated_date = current_date THEN HH_member_array = HH_member_array & HH_member & " "
    	END IF
    Next

    HH_member_array = trim(HH_member_array)
	If HH_member_array <> "" then panels_updated = panels_updated & stat_panel & ","
    HH_member_array = Split(HH_member_array, " ") 	'declaring & splitting the array

	call navigate_to_MAXIS_screen("STAT", "SUMM")
	call autofill_editbox_from_MAXIS(HH_member_array, stat_panel, output_variable)
End Function

'-------------------------------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""		'Connects to BlueZone
call maxis_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
panels_updated = ""

'the dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 141, 70, "Case number dialog"
  EditBox 75, 5, 55, 15, MAXIS_case_number
  EditBox 85, 25, 20, 15, MAXIS_footer_month
  EditBox 110, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 50, 50, 15
    CancelButton 80, 50, 50, 15
  Text 20, 10, 55, 10, "Case Number:"
  Text 20, 30, 65, 10, "Footer month/year:"
EndDialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
  		cancel_without_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		'If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
MAXIS_footer_month_confirmation
current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

'Inputs the panel information for each member/panel if the current date is equal to the update date in MAXIS through custom fumction
'Income panels
Call updated_panel_member_array("JOBS", earned_income)
Call updated_panel_member_array("BUSI", earned_income)
Call updated_panel_member_array("RBIC", earned_income)
Call updated_panel_member_array("UNEA", unearned_income)
'Deductions
Call updated_panel_member_array("COEX", CC_deductions)
Call updated_panel_member_array("DCEX", CC_deductions)
Call updated_panel_member_array("SHEL", housing_costs)
Call updated_panel_member_array("ACUT", housing_costs)

'Special coding for HEST as this is a case based panel
CALL navigate_to_MAXIS_screen("STAT", "HEST")
EMReadScreen updated_date, 8, 21, 55
updated_date = replace(updated_date, " ", "/")

IF updated_date = current_date THEN
	STATS_COUNTER = STATS_counter + 1                      'adds one instance to the stats counter
	call autofill_editbox_from_MAXIS("", "HEST", housing_costs)
	panels_updated = panels_updated & "HEST,"
End if

'Other changes
Call updated_panel_member_array("DISA", DISA_PBEN)
Call updated_panel_member_array("PBEN", DISA_PBEN)
Call updated_panel_member_array("IMIG", imig_info)
Call updated_panel_member_array("WREG", WREG_info)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 321, 235, "Information updated in MAXIS by QI for Case #: " & MAXIS_case_number
  EditBox 70, 195, 245, 15, other_notes
  EditBox 70, 215, 145, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 215, 45, 15
    CancelButton 270, 215, 45, 15
  EditBox 80, 20, 230, 15, earned_income
  EditBox 80, 40, 230, 15, unearned_income
  EditBox 80, 75, 230, 15, CC_deductions
  EditBox 80, 95, 230, 15, housing_costs
  EditBox 80, 130, 230, 15, DISA_PBEN
  EditBox 80, 150, 230, 15, imig_info
  EditBox 80, 170, 230, 15, WREG_info
  Text 30, 80, 45, 10, "COEX/DCEX:"
  Text 55, 155, 20, 10, "IMIG:"
  Text 50, 175, 25, 10, "WREG:"
  Text 55, 45, 25, 10, "UNEA:"
  Text 35, 135, 45, 10, "DISA/PBEN:"
  Text 10, 100, 70, 10, "SHEL/HEST/ACUT:"
  Text 25, 200, 40, 10, "Other notes:"
  Text 15, 25, 60, 10, "JOBS/BUSI/RBIC:"
  Text 10, 220, 60, 10, "Worker signature: "
  GroupBox 5, 10, 310, 50, "Income:"
  GroupBox 5, 65, 310, 50, "Deductions:"
  GroupBox 5, 120, 310, 70, "Other updates:"
EndDialog

'the main dialog
Do
	Do
  		err_msg = ""
  		Dialog Dialog1
  		cancel_confirmation
  		If worker_signature = "" then err_msg = err_msg & vbnewline & "* Enter your worker signature."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------THE CASE NOTE
renewal_period = MAXIS_footer_month & "/" & MAXIS_footer_year		'establishing the renewal period for the header of the case note

panels_updated = trim(panels_updated)								'DO loop removes the excess spaces and commas at the end of the string
Do
	If right(panels_updated, 1) = "," then panels_updated = left(panels_updated, len(panels_updated) - 1)
	panels_updated = trim(panels_updated)
Loop until right(panels_updated, 1) <> ","

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("*" & renewal_period & " recert accuracy update for: " & panels_updated & "*")
Call write_variable_in_CASE_NOTE("This case info has been reviewed and updated by the Quality Improvement team. Do not update the following info unless a new change has been reported:")
If earned_income <> "" or unearned_income <> "" then
	Call write_variable_in_CASE_NOTE("--Income--")
	Call write_bullet_and_variable_in_CASE_NOTE("JOBS/BUSI/RBIC", earned_income)
	Call write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
End if
If housing_costs <> "" or CC_deductions <> "" then
	Call write_variable_in_CASE_NOTE("--Deductions--")
	Call write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", CC_deductions)
	Call write_bullet_and_variable_in_CASE_NOTE("SHEL/HEST/ACUT", housing_costs)
End if
If DISA_PBEN <> "" or imig_info <> "" or WREG_info <> "" then
	Call write_variable_in_CASE_NOTE("--Other updates--")
	Call write_bullet_and_variable_in_CASE_NOTE("DISA/PBEN", DISA_PBEN)
	Call write_bullet_and_variable_in_CASE_NOTE("IMIG", imig_info)
	Call write_bullet_and_variable_in_CASE_NOTE("WREG", WREG_info)
End if
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! Script run complete.")
