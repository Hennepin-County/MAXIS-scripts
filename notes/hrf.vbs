'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HRF.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
Call changelog_update("02/03/2025", "Support for MFIP, GA, and UHFS Budgeting Workaround to support the policy change to eliminate Monthly Reporting and Retrospective Budgeting.##~## ##~##These updates follow the Guide to Six-Month Budgeting available in SIR.##~## ##~##As with any new functionality, but particularly when the supporting policy is also new, reach out with any questions or script errors", "Mark Riegel, Hennepin County")
Call changelog_update("03/27/2024", "Added a checkbox option to indicate that a future month HRF has not been received when processing a HRF for the current month. This adds a line to the CASE/NOTE indicating this future HRF is not received.", "Casey Love, Hennepin County")
Call changelog_update("02/27/2024", "Removed eligibility details from case note. Please use NOTES-Eligibility Summary to document this information.", "Megan Geissler, Hennepin County")
Call changelog_update("06/26/2023", "Added handling to support selection of specific programs for HRF processing.", "Ilse Ferris, Hennepin County")
Call changelog_update("07/10/2019", "Fixed a bug that prevented the script from reading the grant amount if Significant Change was applied on MFIP. Additionally added functionality to copy significant change information into the casenote if ELIG/MF is read.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
call changelog_update("04/23/2018", "Added NOTES on INCOME field and some preselected options to input on NOTES on INCOME field for more detailed case notes.", "Casey Love, Hennepin County")
call changelog_update("02/23/2018", "Added closing message to reminder to workers to accept all work items upon processing HRF's.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2016", "Added seperate functionality for LTC HRF cases.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect "" 'Connecting to BlueZone

' Functions
function HRF_autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
	'--- This function autofills information for all HH members idenified from the HH_member_array from a selected MAXIS panel into an edit box in a dialog.
	'~~~~~ HH_member_array: array of HH members from function HH_member_custom_dialog(HH_member_array). User selects which HH members are added to array.
	'~~~~~ read_panel_from: first four characters because we use separate handling for HCRE-retro. This is something that should be fixed someday!!!!!!!!!
	'~~~~~ variable_written_to: the variable used by the editbox you wish to autofill.
	'===== Keywords: MAXIS, autofill, HH_member_array
	call navigate_to_MAXIS_screen("STAT", left(panel_read_from, 4))

	'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
	EMReadScreen panel_total_check, 6, 2, 73
	IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

	If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
	If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
		EMReadScreen ABPS_total_pages, 1, 2, 78
		If ABPS_total_pages <> 0 then
		Do
			'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
			EMReadScreen support_coop_code, 1, 4, 73
			If support_coop_code = "N" then
			EMReadScreen caregiver_ref_nbr, 2, 4, 47
			If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated.
			End if
			'Then it gets info on the ABPS themself.
			EMReadScreen ABPS_current, 45, 10, 30
			If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
			ABPS_current = replace(ABPS_current, "  First:", ",")
			ABPS_current = replace(ABPS_current, "_", "")
			ABPS_current = trim(ABPS_current)
			ABPS_current = split(ABPS_current)
			For each ABPS_part in ABPS_current
				If ABPS_part <> "" Then
					first_letter = ucase(left(ABPS_part, 1))
					other_letters = LCase(right(ABPS_part, len(ABPS_part) -1))
					If len(ABPS_part) > 1 then
						new_ABPS_current = new_ABPS_current & first_letter & other_letters & " "
					Else
						new_ABPS_current = new_ABPS_current & ABPS_part & " "
					End if
				End If
			Next
			ABPS_row = 15 'Setting variable for do...loop
			Do
			Do 'Using a do...loop to determine which MEMB numbers are with this parent
				EMReadScreen child_ref_nbr, 2, ABPS_row, 35
				If child_ref_nbr <> "__" then
				amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
				children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
				End if
				ABPS_row = ABPS_row + 1
			Loop until ABPS_row > 17		'End of the row
			EMReadScreen more_check, 7, 19, 66
			If more_check = "More: +" then
				EMSendKey "<PF20>"
				EMWaitReady 0, 0
				ABPS_row = 15
			End if
			Loop until more_check <> "More: +"
			'Cleaning up the "children_for_ABPS" variable to be more readable
			If children_for_ABPS = "" Then
				stop_message = "The script you are running " & replace(name_of_script, ".vbs", "") & " is attempting to read information from ABPS. This ABPS panel does not have any children listed. Review the STAT panels, particularly about parental relationships (ABPS/PARE). This panel needs update, or may need to be deleted."
				script_end_procedure(stop_message)
			End If
			children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
			children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
			children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
			children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around
			if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
			if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
			variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
			'Resetting variables for the do...loop in case this function runs again
			new_ABPS_current = ""
			amt_of_children_for_ABPS = 0
			children_for_ABPS = ""
			'Checking to see if it needs to run again, if it does it transmits or else the loop stops
			EMReadScreen ABPS_current_page, 1, 2, 73
			If ABPS_current_page <> ABPS_total_pages then transmit
		Loop until ABPS_current_page = ABPS_total_pages
		'Combining the two variables (support coop and the variable written to)
		variable_written_to = support_coop & variable_written_to
		End if
	Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen ACCI_total, 1, 2, 78
		If ACCI_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_ACCI_to_variable(variable_written_to)
			EMReadScreen ACCI_panel_current, 1, 2, 73
			If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
			Loop until cint(ACCI_panel_current) = cint(ACCI_total)
		End if
		Next
	Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen ACCT_total, 2, 2, 78
		ACCT_total = trim(ACCT_total)   'deleting space if one digit.
		If ACCT_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_ACCT_to_variable(variable_written_to)
			EMReadScreen ACCT_panel_current, 2, 2, 72
			ACCT_panel_current = trim(ACCT_panel_current)
			If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
			Loop until cint(ACCT_panel_current) = cint(ACCT_total)
		End if
		Next
	ElseIf panel_read_from = "ACUT" Then '----------------------------------------------------------------------------------------------------ACUT
		For each HH_member in HH_member_array
			EMWriteScreen HH_member, 20, 76
			transmit
			EMReadScreen ACUT_total, 1, 2, 78
			If ACUT_total <> 0 then
				EMReadScreen share_yn, 1, 6, 42
				EMReadScreen retro_heat_verif, 1, 10, 35
				EMReadScreen retro_heat_amount, 8, 10, 41
				EMReadScreen retro_air_verif, 1, 11, 35
				EMReadScreen retro_air_amount, 8, 11, 41
				EMReadScreen retro_elec_verif, 1, 12, 35
				EMReadScreen retro_elec_amount, 8, 12, 41
				EMReadScreen retro_fuel_verif, 1, 13, 35
				EMReadScreen retro_fuel_amount, 8, 13, 41
				EMReadScreen retro_garbage_verif, 1, 14, 35
				EMReadScreen retro_garbage_amount, 8, 14, 41
				EMReadScreen retro_water_verif, 1, 15, 35
				EMReadScreen retro_water_amount, 8, 15, 41
				EMReadScreen retro_sewer_verif, 1, 16, 35
				EMReadScreen retro_sewer_amount, 8, 16, 41
				EMReadScreen retro_other_verif, 1, 17, 35
				EMReadScreen retro_other_amount, 8, 17, 41

				EMReadScreen prosp_heat_verif, 1, 10, 55
				EMReadScreen prosp_heat_amount, 8, 10, 61
				EMReadScreen prosp_air_verif, 1, 11, 55
				EMReadScreen prosp_air_amount, 8, 11, 61
				EMReadScreen prosp_elec_verif, 1, 12, 55
				EMReadScreen prosp_elec_amount, 8, 12, 61
				EMReadScreen prosp_fuel_verif, 1, 13, 55
				EMReadScreen prosp_fuel_amount, 8, 13, 61
				EMReadScreen prosp_garbage_verif, 1, 14, 55
				EMReadScreen prosp_garbage_amount, 8, 14, 61
				EMReadScreen prosp_water_verif, 1, 15, 55
				EMReadScreen prosp_water_amount, 8, 15, 61
				EMReadScreen prosp_sewer_verif, 1, 16, 55
				EMReadScreen prosp_sewer_amount, 8, 16, 61
				EMReadScreen prosp_other_verif, 1, 17, 55
				EMReadScreen prosp_other_amount, 8, 17, 61

				EMReadScreen dwp_phone_yn, 1, 18, 55

				variable_written_to = "Actutal Utilitiy Expense for M" & HH_member
				If share_yn = "Y" Then variable_written_to = variable_written_to & " - this expense is shared."
				If retro_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (retro) $" & trim(retro_heat_amount) & " - Verif: " & retro_heat_verif & "."
				If prosp_heat_verif <> "_" Then variable_written_to = variable_written_to & " Heat (prosp) $" & trim(prosp_heat_amount) & " - Verif: " & prosp_heat_verif & "."
				If retro_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (retro) $" & trim(retro_air_amount) & " - Verif: " & retro_air_verif & "."
				If prosp_air_verif <> "_" Then variable_written_to = variable_written_to & " Air (prosp) $" & trim(prosp_air_amount) & " - Verif: " & prosp_air_verif & "."
				If retro_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (retro) $" & trim(retro_elec_amount) & " - Verif: " & retro_elec_verif & "."
				If prosp_elec_verif <> "_" Then variable_written_to = variable_written_to & " Electric (prosp) $" & trim(prosp_elec_amount) & " - Verif: " & prosp_elec_verif & "."
				If retro_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (retro) $" & trim(retro_fuel_amount) & " - Verif: " & retro_fuel_verif & "."
				If prosp_fuel_verif <> "_" Then variable_written_to = variable_written_to & " Fuel (prosp) $" & trim(prosp_fuel_amount) & " - Verif: " & prosp_fuel_verif & "."
				If retro_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (retro) $" & trim(retro_garbage_amount) & " - Verif: " & retro_garbage_verif & "."
				If prosp_garbage_verif <> "_" Then variable_written_to = variable_written_to & " Garbage (prosp) $" & trim(prosp_garbage_amount) & " - Verif: " & prosp_garbage_verif & "."
				If retro_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (retro) $" & trim(retro_water_amount) & " - Verif: " & retro_water_verif & "."
				If prosp_water_verif <> "_" Then variable_written_to = variable_written_to & " Water (prosp) $" & trim(prosp_water_amount) & " - Verif: " & prosp_water_verif & "."
				If retro_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (retro) $" & trim(retro_sewer_amount) & " - Verif: " & retro_sewer_verif & "."
				If prosp_sewer_verif <> "_" Then variable_written_to = variable_written_to & " Sewer (prosp) $" & trim(prosp_sewer_amount) & " - Verif: " & prosp_sewer_verif & "."
				If retro_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (retro) $" & trim(retro_other_amount) & " - Verif: " & retro_other_verif & "."
				If prosp_other_verif <> "_" Then variable_written_to = variable_written_to & " Other (prosp) $" & trim(prosp_other_amount) & " - Verif: " & prosp_other_verif & "."
				If dwp_phone_yn = "Y" Then variable_written_to = variable_written_to & " Standard DWP Phone allowance of $35."

			End If
		Next
	Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
		EMReadScreen addr_line_01, 22, 6, 43
		EMReadScreen addr_line_02, 22, 7, 43
		EMReadScreen city_line, 15, 8, 43
		EMReadScreen state_line, 2, 8, 66
		EMReadScreen zip_line, 12, 9, 43
		variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
		variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
	Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
		EMReadScreen AREP_name, 37, 4, 32
		AREP_name = replace(AREP_name, "_", "")
		AREP_name = split(AREP_name)
		For each word in AREP_name
		If word <> "" then
			first_letter_of_word = ucase(left(word, 1))
			rest_of_word = LCase(right(word, len(word) -1))
			If len(word) > 2 then
			variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
			Else
			variable_written_to = variable_written_to & word & " "
			End if
		End if
		Next
	Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
		EMReadScreen BILS_amt, 1, 2, 78
		If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
	Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen BUSI_total, 1, 2, 78
		If BUSI_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call HRF_add_BUSI_to_variable(variable_written_to)
			EMReadScreen BUSI_panel_current, 1, 2, 73
			If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
			Loop until cint(BUSI_panel_current) = cint(BUSI_total)
		End if
		Next
	Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen CARS_total, 2, 2, 78
		CARS_total = trim(CARS_total)
		If CARS_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_CARS_to_variable(variable_written_to)
			EMReadScreen CARS_panel_current, 2, 2, 72
			CARS_panel_current = trim(CARS_panel_current)
			If cint(CARS_panel_current) < cint(CARS_total) then transmit
			Loop until cint(CARS_panel_current) = cint(CARS_total)
		End if
		Next
	Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen cash_amt, 8, 8, 39
		cash_amt = trim(cash_amt)
		If cash_amt <> "________" then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
		End if
		Next
	Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen support_amt, 8, 10, 63
		support_amt = trim(support_amt)
		If support_amt <> "________" then
			EMReadScreen support_ver, 1, 10, 36
			If support_ver = "?" or support_ver = "N" then
			support_ver = ", no proof provided"
			Else
			support_ver = ""
			End if
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
		End if
		EMReadScreen alimony_amt, 8, 11, 63
		alimony_amt = trim(alimony_amt)
		If alimony_amt <> "________" then
			EMReadScreen alimony_ver, 1, 11, 36
			If alimony_ver = "?" or alimony_ver = "N" then
			alimony_ver = ", no proof provided"
			Else
			alimony_ver = ""
			End if
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
		End if
		EMReadScreen tax_dep_amt, 8, 12, 63
		tax_dep_amt = trim(tax_dep_amt)
		If tax_dep_amt <> "________" then
			EMReadScreen tax_dep_ver, 1, 12, 36
			If tax_dep_ver = "?" or tax_dep_ver = "N" then
			tax_dep_ver = ", no proof provided"
			Else
			tax_dep_ver = ""
			End if
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
		End if
		EMReadScreen other_COEX_amt, 8, 13, 63
		other_COEX_amt = trim(other_COEX_amt)
		If other_COEX_amt <> "________" then
			EMReadScreen other_COEX_ver, 1, 13, 36
			If other_COEX_ver = "?" or other_COEX_ver = "N" then
			other_COEX_ver = ", no proof provided"
			Else
			other_COEX_ver = ""
			End if
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
		End if
		Next
	Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen DCEX_total, 1, 2, 78
		If DCEX_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
				DCEX_row = 11
				Do
					EMReadScreen expense_amt, 8, DCEX_row, 63
					expense_amt = trim(expense_amt)
					If expense_amt <> "________" then
						EMReadScreen child_ref_nbr, 2, DCEX_row, 29
						EMReadScreen expense_ver, 1, DCEX_row, 41
						If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
							expense_ver = ", no proof provided"
						Else
							expense_ver = ""
						End if
						variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
					End if
					DCEX_row = DCEX_row + 1
				Loop until DCEX_row = 17
				EMReadScreen DCEX_panel_current, 1, 2, 73
				If cint(DCEX_panel_current) < cint(DCEX_total) then transmit
			Loop until cint(DCEX_panel_current) = cint(DCEX_total)
		End if
		Next
	Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		DIET_row = 8 'Setting this variable for the next do...loop
		EMReadScreen DIET_total, 1, 2, 78
		If DIET_total <> 0 then
			DIET = DIET & "Member " & HH_member & "- "
			Do
			EMReadScreen diet_type, 2, DIET_row, 40
			EMReadScreen diet_proof, 1, DIET_row, 51
			If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then
				diet_proof = ", no proof provided"
			Else
				diet_proof = ""
			End if
			If diet_type = "01" then diet_type = "High Protein"
			If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
			If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
			If diet_type = "04" then diet_type = "Lo Cholesterol"
			If diet_type = "05" then diet_type = "High Residue"
			If diet_type = "06" then diet_type = "Preg/Lactation"
			If diet_type = "07" then diet_type = "Gluten Free"
			If diet_type = "08" then diet_type = "Lactose Free"
			If diet_type = "09" then diet_type = "Anti-Dumping"
			If diet_type = "10" then diet_type = "Hypoglycemic"
			If diet_type = "11" then diet_type = "Ketogenic"
			If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
			DIET_row = DIET_row + 1
			Loop until DIET_row = 19
		End if
		Next
	Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadscreen DISA_total, 1, 2, 78
		IF DISA_total <> 0 THEN
			'Reads and formats CASH/GRH disa status
			EMReadScreen CASH_DISA_status, 2, 11, 59
			EMReadScreen CASH_DISA_verif, 1, 11, 69
			IF CASH_DISA_status = "01" or CASH_DISA_status = "02" or CASH_DISA_status = "03" OR CASH_DISA_status = "04" THEN CASH_DISA_status = "RSDI/SSI certified"
			IF CASH_DISA_status = "06" THEN CASH_DISA_status = "SMRT/SSA pends"
			IF CASH_DISA_status = "08" THEN CASH_DISA_status = "Certified Blind"
			IF CASH_DISA_status = "09" THEN CASH_DISA_status = "Ill/Incap"
			IF CASH_DISA_status = "10" THEN CASH_DISA_status = "Certified disabled"
			IF CASH_DISA_verif = "?" OR CASH_DISA_verif = "N" THEN
				CASH_DISA_verif = ", no proof provided"
			ELSE
				CASH_DISA_verif = ""
			END IF

			'Reads and formats SNAP disa status
			EmreadScreen SNAP_DISA_status, 2, 12, 59
			EMReadScreen SNAP_DISA_verif, 1, 12, 69
			IF SNAP_DISA_status = "01" or SNAP_DISA_status = "02" or SNAP_DISA_status = "03" OR SNAP_DISA_status = "04" THEN SNAP_DISA_status = "RSDI/SSI certified"
			IF SNAP_DISA_status = "08" THEN SNAP_DISA_status = "Certified Blind"
			IF SNAP_DISA_status = "09" THEN SNAP_DISA_status = "Ill/Incap"
			IF SNAP_DISA_status = "10" THEN SNAP_DISA_status = "Certified disabled"
			IF SNAP_DISA_status = "11" THEN SNAP_DISA_status = "VA determined PD disa"
			IF SNAP_DISA_status = "12" THEN SNAP_DISA_status = "VA (other accept disa)"
			IF SNAP_DISA_status = "13" THEN SNAP_DISA_status = "Cert RR Ret Disa & on MEDI"
			IF SNAP_DISA_status = "14" THEN SNAP_DISA_status = "Other Govt Perm Disa Ret Bnft"
			IF SNAP_DISA_status = "15" THEN SNAP_DISA_status = "Disability from MINE list"
			IF SNAP_DISA_status = "16" THEN SNAP_DISA_status = "Unable to p&p own meal"
			IF SNAP_DISA_verif = "?" OR SNAP_DISA_verif = "N" THEN
				SNAP_DISA_verif = ", no proof provided"
			ELSE
				SNAP_DISA_verif = ""
			END IF

			'Reads and formats HC disa status/verif
			EMReadScreen HC_DISA_status, 2, 13, 59
			EMReadScreen HC_DISA_verif, 1, 13, 69
			If HC_DISA_status = "01" or HC_DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
			If HC_DISA_status = "06" then HC_DISA_status = "SMRT/SSA pends"
			If HC_DISA_status = "08" then HC_DISA_status = "Certified blind"
			If HC_DISA_status = "10" then HC_DISA_status = "Certified disabled"
			If HC_DISA_status = "11" then HC_DISA_status = "Spec cat- disa child"
			If HC_DISA_status = "20" then HC_DISA_status = "TEFRA- disabled"
			If HC_DISA_status = "21" then HC_DISA_status = "TEFRA- blind"
			If HC_DISA_status = "22" then HC_DISA_status = "MA-EPD"
			If HC_DISA_status = "23" then HC_DISA_status = "MA/waiver"
			If HC_DISA_status = "24" then HC_DISA_status = "SSA/SMRT appeal pends"
			If HC_DISA_status = "26" then HC_DISA_status = "SSA/SMRT disa deny"
			IF HC_DISA_verif = "?" OR HC_DISA_verif = "N" THEN
				HC_DISA_verif = ", no proof provided"
			ELSE
				HC_DISA_verif = ""
			END IF
			'cleaning to make variable to write
			IF CASH_DISA_status = "__" THEN
				CASH_DISA_status = ""
			ELSE
				IF CASH_DISA_status = SNAP_DISA_status THEN
					SNAP_DISA_status = "__"
					CASH_DISA_status = "CASH/SNAP: " & CASH_DISA_status & " "
				ELSE
					CASH_DISA_status = "CASH: " & CASH_DISA_status & " "
				END IF
			END IF
			IF SNAP_DISA_status = "__" THEN
				SNAP_DISA_status = ""
			ELSE
				SNAP_DISA_status = "SNAP: " & SNAP_DISA_status & " "
			END IF
			IF HC_DISA_status = "__" THEN
				HC_DISA_status = ""
			ELSE
				HC_DISA_status = "HC: " & HC_DISA_status & " "
			END IF
			'Adding verif code info if N or ?
			IF CASH_DISA_verif <> "" THEN CASH_DISA_status = CASH_DISA_status & CASH_DISA_verif & " "
			IF SNAP_DISA_verif <> "" THEN SNAP_DISA_status = SNAP_DISA_status & SNAP_DISA_verif & " "
			IF HC_DISA_verif <> "" THEN HC_DISA_status = HC_DISA_status & HC_DISA_verif & " "
			'Creating final variable
			IF CASH_DISA_status <> "" THEN FINAL_DISA_status = CASH_DISA_status
			IF SNAP_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & SNAP_DISA_status
			IF HC_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & HC_DISA_status

			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & FINAL_DISA_status & "; "
		END IF
		Next
	Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
		row = 14
		Do
		EMReadScreen reference_numbers_current_row, 40, row, 39
		reference_numbers = reference_numbers + reference_numbers_current_row
		row = row + 1
		Loop until row = 18
		reference_numbers = replace(reference_numbers, "  ", " ")
		reference_numbers = split(reference_numbers)
		For each member in reference_numbers
		If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
		Next
		EATS_info = trim(EATS_info)
		if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
		If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
	Elseif panel_read_from = "EMPS" then '----------------------------------------------------------------------------------------------------EMPS
			For each HH_member in HH_member_array
			'blanking out variables for the next HH member
			EMPS_info = ""
			ES_exemptions = ""
			ES_info = ""
			EMWriteScreen HH_member, 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen EMPS_total, 1, 2, 78
			If EMPS_total <> 0 then
				'orientation info (EMPS_info variable)-------------------------------------------------------------------------
				EMReadScreen EMPS_orientation_date, 8, 5, 39
				IF EMPS_orientation_date = "__ __ __" then
					EMPS_orientation_date = "none"
				ElseIf EMPS_orientation_date <> "__ __ __" then
					EMPS_orientation_date = replace(EMPS_orientation_date, " ", "/")
					EMPS_info = EMPS_info & " Fin orient: " & EMPS_orientation_date & ","
				END IF
				EMReadScreen EMPS_orientation_attended, 1, 5, 65
				IF EMPS_orientation_attended <> "_" then EMPS_info = EMPS_info & " Attended orient: " & EMPS_orientation_attended & ","
				'Good cause (EMPS_info variable)
				EMReadScreen EMPS_good_cause, 2, 5, 79
				IF EMPS_good_cause <> "__" then
					If EMPS_good_cause = "01" then EMPS_good_cause = "01-No Good Cause"
					If EMPS_good_cause = "02" then EMPS_good_cause = "02-No Child Care"
					If EMPS_good_cause = "03" then EMPS_good_cause = "03-Ill or Injured"
					If EMPS_good_cause = "04" then EMPS_good_cause = "04-Care Ill/Incap. Family Member"
					If EMPS_good_cause = "05" then EMPS_good_cause = "05-Lack of Transportation"
					If EMPS_good_cause = "06" then EMPS_good_cause = "06-Emergency"
					If EMPS_good_cause = "07" then EMPS_good_cause = "07-Judicial Proceedings"
					If EMPS_good_cause = "08" then EMPS_good_cause = "08-Conflicts with Work/School"
					If EMPS_good_cause = "09" then EMPS_good_cause = "09-Other Impediments"
					If EMPS_good_cause = "10" then EMPS_good_cause = "10-Special Medical Criteria "
					If EMPS_good_cause = "20" then EMPS_good_cause = "20-Exempt--Only/1st Caregiver Employed 35+ Hours"
					If EMPS_good_cause = "21" then EMPS_good_cause = "21-Exempt--2nd Caregiver Employed 20+ Hours"
					If EMPS_good_cause = "22" then EMPS_good_cause = "22-Exempt--Preg/Parenting Caregiver < Age 20"
					If EMPS_good_cause = "23" then EMPS_good_cause = "23-Exempt--Special Medical Criteria"
					IF EMPS_good_cause <> "__" then EMPS_info = EMPS_info & " Good cause: " & EMPS_good_cause & ","
				END IF

				'sanction dates (EMPS_info variable)
				EMReadScreen EMPS_sanc_begin, 8, 6, 39
				If EMPS_sanc_begin <> "__ 01 __" then
					EMPS_sanc_begin = replace(EMPS_sanc_begin, "_", "/")
					sanction_date = sanction_date & EMPS_sanc_begin
				END IF
				EMReadScreen EMPS_sanc_end, 8, 6, 65
				If EMPS_sanc_end <> "__ 01 __" then
					EMPS_sanc_end = replace(EMPS_sanc_end, "_", "/")
					sanction_date = sanction_date & "-" & EMPS_sanc_end
				END IF
				IF sanction_date <> "" then EMPS_info = EMPS_info & " Sanction dates: " & sanction_date & ","
				'cleaning up ES_info variable
				If right(EMPS_info, 1) = "," then EMPS_info = left(EMPS_info, len(EMPS_info) - 1)
				IF trim(EMPS_info) <> "" then EMPS_info = EMPS_info & "."

				'other sanction dates (ES_exemptions variable)--------------------------------------------------------------------------------
				'special medical criteria
				EMReadScreen EMPS_memb_at_home, 1, 8, 76
				IF EMPS_memb_at_home <> "N" then
					If EMPS_memb_at_home = "1" then EMPS_memb_at_home = "Home-Health/Waiver service"
					IF EMPS_memb_at_home = "2" then EMPS_memb_at_home = "Child w/ severe emotional dist"
					IF EMPS_memb_at_home = "3" then EMPS_memb_at_home = "Adult/Serious Persistent MI"
					ES_exemptions = ES_exemptions & " Special med criteria: " & EMPS_memb_at_home & ","
				END IF

				EMReadScreen EMPS_care_family, 1, 9, 76
				IF EMPS_care_family = "Y" then ES_exemptions = ES_exemptions & " Care of ill/incap memb: " & EMPS_care_family & ","
				EMReadScreen EMPS_crisis, 1, 10, 76
				IF EMPS_crisis = "Y" then ES_exemptions = ES_exemptions & " Family crisis: " & EMPS_crisis & ","

				'hard to employ
				EMReadScreen EMPS_hard_employ, 2, 11, 76
				IF EMPS_hard_employ <> "NO" then
					IF EMPS_hard_employ = "IQ" then EMPS_hard_employ = "IQ tested at < 80"
					IF EMPS_hard_employ = "LD" then EMPS_hard_employ = "Learning Disabled"
					IF EMPS_hard_employ = "MI" then EMPS_hard_employ = "Mentally ill"
					IF EMPS_hard_employ = "DD" then EMPS_hard_employ = "Dev Disabled"
					IF EMPS_hard_employ = "UN" then EMPS_hard_employ = "Unemployable"
					ES_exemptions = ES_exemptions & " Hard to employ: " & EMPS_hard_employ & ","
				END IF

				'EMPS under 1 coding and dates used(ES_exemptions variable)
				EMReadScreen EMPS_under1, 1, 12, 76
				IF EMPS_under1 = "Y" then
					ES_exemptions = ES_exemptions & " FT child under 1: " & EMPS_under1 & ","
					EMWriteScreen "X", 12, 39
					transmit
					MAXIS_row = 7
					MAXIS_col = 22
					DO
						EMReadScreen exemption_date, 9, MAXIS_row, MAXIS_col
						If trim(exemption_date) = "" then exit do
						If exemption_date <> "__ / ____" then
							child_under1_dates = child_under1_dates & exemption_date & ", "
							MAXIS_col = MAXIS_col + 11
							If MAXIS_col = 66 then
								MAXIS_row = MAXIS_row + 1
								MAXIS_col = 22
							END IF
						END IF
					LOOP until exemption_date = "__ / ____" or (MAXIS_row = 9 and MAXIS_col = 66)
					PF3
					'cleaning up excess comma at the end of child_under1_dates variable
					If right(child_under1_dates,  2) = ", " then child_under1_dates = left(child_under1_dates, len(child_under1_dates) - 2)
					If trim(child_under1_dates) = "" then child_under1_dates = " N/A"
					ES_exemptions = ES_exemptions & " Child under 1 exeption dates: " & child_under1_dates & ","
				END IF

				'cleaning up ES_exemptions variable
				If right(ES_exemptions, 1) = "," then ES_exemptions = left(ES_exemptions, len(ES_exemptions) - 1)
				IF trim(ES_exemptions) <> "" then ES_exemptions = ES_exemptions & "."

				'Reading ES Information (for ES_info variable)
				EMReadScreen ES_status, 40, 15, 40
				ES_status = trim(ES_status)
				IF ES_status <> "" then ES_info = ES_info & " ES status: " & ES_status & ","
				EMReadScreen ES_referral_date, 8, 16, 40
				If ES_referral_date <> "__ __ __" then
					ES_referral_date = replace(ES_referral_date, " ", "/")
					ES_info = ES_info & " ES referral date: " & ES_referral_date & ","
				END IF

				EMReadScreen DWP_plan_date, 8, 17, 40
				IF DWP_plan_date <> "__ __ __" then
					DWP_plan_date = replace(DWP_plan_date, "_", "/")
					ES_info = ES_info & " DWP plan date: " & DWP_plan_date & ","
				END IF

				EMReadScreen minor_ES_option, 2, 16, 76
				If minor_ES_option <> "__" then
					IF minor_ES_option = "SC" then minor_ES_option = "Secondary Education"
					IF minor_ES_option = "EM" then minor_ES_option = "Employment"
					ES_info = ES_info & " 18/19 yr old ES option: " & minor_ES_option & ","
				END if

				'cleaning up ES_info variable
				If right(ES_info, 1) = "," then ES_info = left(ES_info, len(ES_info) - 1)

				variable_written_to = variable_written_to & "Member " & HH_member & "- "
				variable_written_to = variable_written_to & EMPS_info & ES_exemptions & ES_info & "; "
			END IF
		next
	Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen FACI_total, 1, 2, 78
		If FACI_total <> 0 then
			row = 14
			Do
			EMReadScreen date_in_check, 4, row, 53
			EMReadScreen date_in_month_day, 5, row, 47
			EMReadScreen date_out_check, 4, row, 77
			date_in_month_day = replace(date_in_month_day, " ", "/") & "/"
			If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
			If row > 18 then
				EMReadScreen FACI_page, 1, 2, 73
				If FACI_page = FACI_total then
				FACI_status = "Not in facility"
				Else
				transmit
				row = 14
				End if
			End if
			Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
			EMReadScreen client_FACI, 30, 6, 43
			client_FACI = replace(client_FACI, "_", "")
			FACI_array = split(client_FACI)
			For each a in FACI_array
			If a <> "" then
				b = ucase(left(a, 1))
				c = LCase(right(a, len(a) -1))
				new_FACI = new_FACI & b & c & " "
			End if
			Next
			client_FACI = new_FACI
			If FACI_status = "Not in facility" then
			client_FACI = ""
			Else
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			variable_written_to = variable_written_to & client_FACI & " Date in: " & date_in_month_day & date_in_check & "; "
			End if
		End if
		Next
	Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
		For each HH_member in HH_member_array
		EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
		If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		fmed_row = 9 'Setting this variable for the next do...loop
		EMReadScreen fmed_total, 1, 2, 78
		If fmed_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			use_expense = False					'<--- Used to determine if an FMED expense that has an end date is going to be counted.
			EMReadScreen fmed_type, 2, fmed_row, 25
			EMReadScreen fmed_proof, 2, fmed_row, 32
			EMReadScreen fmed_amt, 8, fmed_row, 70
			EMReadScreen fmed_end_date, 5, fmed_row, 60		'reading end date to see if this one even gets added.
			IF fmed_end_date <> "__ __" THEN
				fmed_end_date = replace(fmed_end_date, " ", "/01/")
				fmed_end_date = dateadd("M", 1, fmed_end_date)
				fmed_end_date = dateadd("D", -1, fmed_end_date)
				IF datediff("D", date, fmed_end_date) > 0 THEN use_expense = True		'<--- If the end date of the FMED expense is the current month or a future month, the expense is going to be counted.
			END IF
			If fmed_end_date = "__ __" OR use_expense = TRUE then					'Skips entries with an end date or end dates in the past.
				If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then
				fmed_proof = ", no proof provided"
				Else
				fmed_proof = ""
				End if
				If fmed_amt = "________" then
				fmed_amt = ""
				Else
				fmed_amt = " ($" & trim(fmed_amt) & ")"
				End if
				If fmed_type = "01" then fmed_type = "Nursing Home"
				If fmed_type = "02" then fmed_type = "Hosp/Clinic"
				If fmed_type = "03" then fmed_type = "Physicians"
				If fmed_type = "04" then fmed_type = "Prescriptions"
				If fmed_type = "05" then fmed_type = "Ins Premiums"
				If fmed_type = "06" then fmed_type = "Dental"
				If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
				If fmed_type = "08" then fmed_type = "Vision Care"
				If fmed_type = "09" then fmed_type = "Medicare Prem"
				If fmed_type = "10" then fmed_type = "Mo. Spdwn Amt/Waiver Obl"
				If fmed_type = "11" then fmed_type = "Home Care"
				If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
				If fmed_type = "15" then fmed_type = "Medi Part D premium"
				If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
				IF fmed_end_date <> "__ __" THEN					'<--- If there is a counted FMED expense with a future end date, the script will modify the way that end date is displayed.
					fmed_end_date = datepart("M", fmed_end_date) & "/" & right(datepart("YYYY", fmed_end_date), 2)		'<--- Begins pulling apart fmed_end_date to format it to human speak.
					IF left(fmed_end_date, 1) <> "0" THEN fmed_end_date = "0" & fmed_end_date
					variable_written_to = left(variable_written_to, len(variable_written_to) - 2) & ", counted through " & fmed_end_date & "; "			'<--- Putting variable_written_to back together with FMED expense end date information.
				END IF
			End if
			fmed_row = fmed_row + 1
			If fmed_row = 15 then
				PF20
				fmed_row = 9
				EMReadScreen last_page_check, 21, 24, 2
				If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
			End if
			Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
		End if
		Next
	Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
		EMReadScreen variable_written_to, 8, 10, 51
		variable_written_to = replace(variable_written_to, " ", "/")
		If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
		variable_written_to = replace(variable_written_to, " ", "/")
		If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
		If isdate(variable_written_to) = False then variable_written_to = ""
	Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
		EMReadScreen variable_written_to, 5, 10, 64
		If isdate(variable_written_to) = True then
		variable_written_to = replace(variable_written_to, " ", "/01/")
		If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
			variable_written_to = variable_written_to
		Else
			variable_written_to = ""
		End if
		End if
	Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
		EMReadScreen HEST_total, 1, 2, 78
		If HEST_total <> 0 then
		EMReadScreen heat_air_check, 6, 13, 75
		If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
		EMReadScreen electric_check, 6, 14, 75
		If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
		EMReadScreen phone_check, 6, 15, 75
		If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
		End if
	Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen IMIG_total, 1, 2, 78
		If IMIG_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			EMReadScreen IMIG_type, 30, 6, 48
			variable_written_to = variable_written_to & trim(IMIG_type) & "; "
		End if
		Next
	Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
		EMReadScreen INSA_amt, 1, 2, 78
		If INSA_amt <> 0 then
		'Runs once per INSA screen
			For i = 1 to INSA_amt step 1
				insurance_name = ""
				'Goes to the correct screen
				EMWriteScreen "0" & i, 20, 79
				transmit
				'Gather Insurance Name
				EMReadScreen INSA_name, 38, 10, 38
				INSA_name = replace(INSA_name, "_", "")
				INSA_name = split(INSA_name)
				For each word in INSA_name
					If trim(word) <> "" then
							first_letter_of_word = ucase(left(word, 1))
							rest_of_word = LCase(right(word, len(word) -1))
							If len(word) > 4 then
								insurance_name = insurance_name & first_letter_of_word & rest_of_word & " "
							Else
								insurance_name = insurance_name & word & " "
							End if
					End if
				Next
				'Create a list of members covered by this insurance
				INSA_row = 15 : INSA_col = 30
				insured_count = 0
				member_list = ""
				Do
					EMReadScreen insured_member, 2, INSA_row, INSA_col
					If insured_member <> "__" then
						if member_list = "" then member_list = insured_member
						if member_list <> "" then member_list = member_list & ", " & insured_member
						INSA_col = INSA_col + 4
						If INSA_col = 70 then
							INSA_col = 30 : INSA_row = 16
						End If
					End If
				loop until insured_member = "__"
				'Retain "variable_written_to" as is while also adding members covered by the insurance policy
				'Example - "Members: 01, 03, 07 are covered by Blue Cross Blue Shield; "
				variable_written_to = variable_written_to & "Members: " & member_list & " are covered by " & trim(insurance_name) & "; "
			Next
			'This will loop and add the above statement for all insurance policies listed
		End if
	Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen JOBS_total, 1, 2, 78
		If JOBS_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call HRF_add_JOBS_to_variable(variable_written_to)
			EMReadScreen JOBS_panel_current, 1, 2, 73
			If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
			Loop until cint(JOBS_panel_current) = cint(JOBS_total)
		End if
		Next
	Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen MEDI_amt, 1, 2, 78
		If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
		Next
	Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		transmit
		EMReadScreen rel_to_applicant, 2, 10, 42
		EMReadScreen client_age, 3, 8, 76
		If client_age = "   " then client_age = 0
		If cint(client_age) >= 21 or rel_to_applicant = "02" then
			number_of_adults = number_of_adults + 1
		Else
			number_of_children = number_of_children + 1
		End if
		Next
		If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
		If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
		If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
	Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen citizen, 1, 11, 49
		If citizen = "Y" then citizen = "US citizen"
		If citizen = "N" then citizen = "non-citizen"
		EMReadScreen citizenship_ver, 2, 11, 78
		EMReadScreen SSA_MA_citizenship_ver, 1, 12, 49
		If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
		If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
		If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
		Next
	ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
		EMReadScreen variable_written_to, 8, 6, 39
		variable_written_to = replace(variable_written_to, " ", "/")
		If isdate(variable_written_to) = True then
		variable_written_to = cdate(variable_written_to) & ""
		Else
		variable_written_to = ""
		End if
	Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen OTHR_total, 2, 2, 78
		OTHR_total = trim(OTHR_total)
		If OTHR_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_OTHR_to_variable(variable_written_to)
			EMReadScreen OTHR_panel_current, 2, 2, 72
			OTHR_panel_current = trim(OTHR_panel_current)
			If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
			Loop until cint(OTHR_panel_current) = cint(OTHR_total)
		End if
		Next
	Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		transmit
		EMReadScreen panel_amt, 1, 2, 78
		If panel_amt <> "0" then
			PBEN = PBEN & "Member " & HH_member & "- "
			row = 8
			Do
			EMReadScreen PBEN_type, 12, row, 28
			EMReadScreen PBEN_disp, 1, row, 77
			If PBEN_disp = "A" then PBEN_disp = " appealing"
			If PBEN_disp = "D" then PBEN_disp = " denied"
			If PBEN_disp = "E" then PBEN_disp = " eligible"
			If PBEN_disp = "P" then PBEN_disp = " pends"
			If PBEN_disp = "N" then PBEN_disp = " not applied yet"
			If PBEN_disp = "R" then PBEN_disp = " refused"
			If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
			row = row + 1
			Loop until row = 14
		End if
		Next
		If PBEN <> "" then variable_written_to = variable_written_to & PBEN
	Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen PREG_total, 1, 2, 78
		If PREG_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			EMReadScreen PREG_due_date, 8, 10, 53
			If PREG_due_date = "__ __ __" then
			PREG_due_date = "unknown"
			Else
			PREG_due_date = replace(PREG_due_date, " ", "/")
			End if
			variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
		End if
		Next
	Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
		row = 6
		Do
		EMReadScreen appl_prog_date, 8, row, 33
		If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
		row = row + 1
		Loop until row = 13
		appl_prog_date_array = split(appl_prog_date_array)
		variable_written_to = CDate(appl_prog_date_array(0))
		for i = 0 to ubound(appl_prog_date_array) - 1
		if CDate(appl_prog_date_array(i)) > variable_written_to then
			variable_written_to = CDate(appl_prog_date_array(i))
		End if
		next
		If isdate(variable_written_to) = True then
		variable_written_to = cdate(variable_written_to) & ""
		Else
		variable_written_to = ""
		End if
	Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen RBIC_total, 1, 2, 78
		If RBIC_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_RBIC_to_variable(variable_written_to)
			EMReadScreen RBIC_panel_current, 1, 2, 73
			If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
			Loop until cint(RBIC_panel_current) = cint(RBIC_total)
		End if
		Next
	Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen REST_total, 2, 2, 78
		REST_total = trim(REST_total)
		If REST_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_REST_to_variable(variable_written_to)
			EMReadScreen REST_panel_current, 2, 2, 72
			REST_panel_current = trim(REST_panel_current)
			If cint(REST_panel_current) < cint(REST_total) then transmit
			Loop until cint(REST_panel_current) = cint(REST_total)
		End if
		Next
	Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
		EMReadScreen variable_written_to, 8, 13, 37
		variable_written_to = replace(variable_written_to, " ", "/")
		If isdate(variable_written_to) = True then
		variable_written_to = cdate(variable_written_to) & ""
		Else
		variable_written_to = ""
		End if
	Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
		For each HH_member in HH_member_array
				EMWriteScreen HH_member, 20, 76
				EMWriteScreen "01", 20, 79
				transmit
				EMReadScreen school_type, 2, 7, 40							'Reading the school type code and converting it into words
				If school_type = "01" then school_type = "elementary school"
				If school_type = "11" then school_type = "middle school"
				If school_type = "02" then school_type = "high school"
				If school_type = "03" then school_type = "GED"
				If school_type = "07" then school_type = "IEP"
				If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
				If school_type = "12" then school_type = "adult basic education"
				If school_type = "13" then school_type = "English as a 2nd language"
				If school_type = "06" or school_type = "__" or school_type = "?_" then  'if the school type is blank, child not in school, or postponed default type to blank.
					school_type = ""
				Else
					EMReadScreen SCHL_ver, 2, 6, 63
					If SCHL_ver = "?_" or SCHL_ver = "NO" then								'If the verification field is postponed or NO it defaults to no proof provided
						school_proof_type = ", no proof provided"
					Else
						school_proof_type = ""
					End if
					EMReadScreen FS_eligibility_status_SCHL, 2, 16, 63				'Reading the FS eligibility status and converting it to words
					IF FS_eligibility_status_SCHL = "01" THEN FS_eligibility_status_SCHL = ", FS Elig Status: < 18 or 50+"
					IF FS_eligibility_status_SCHL = "02" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Disabled"
					IF FS_eligibility_status_SCHL = "03" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Attenting Higher Ed or Attending < 1/2"
					IF FS_eligibility_status_SCHL = "04" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Employed 20 Hours/Wk"
					IF FS_eligibility_status_SCHL = "05" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Fed/State Work Study Program"
					IF FS_eligibility_status_SCHL = "06" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant Under 6"
					IF FS_eligibility_status_SCHL = "07" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Dependant 6-11, daycare not available"
					IF FS_eligibility_status_SCHL = "09" THEN FS_eligibility_status_SCHL = ", FS Elig Status: WIA, TAA, TRA, or FSET placement"
					IF FS_eligibility_status_SCHL = "10" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Full Time Single Parent with Child under 12"
					IF FS_eligibility_status_SCHL = "99" THEN FS_eligibility_status_SCHL = ", FS Elig Status: Not Eligible"
					IF FS_eligibility_status_SCHL = "__" or FS_eligibility_status_SCHL = "?_" THEN FS_eligibility_status_SCHL = ""
					'formatting the output variable for the function
					variable_written_to = variable_written_to & "Member " & HH_member & "- "
					variable_written_to = variable_written_to & school_type & school_proof_type & FS_eligibility_status_SCHL & "; "
				End if
			Next
	Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen SECU_total, 2, 2, 78
		SECU_total = trim(SECU_total)
		If SECU_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call add_SECU_to_variable(variable_written_to)
			EMReadScreen SECU_panel_current, 2, 2, 72
			SECU_panel_current = trim(SECU_panel_current)
			If cint(SECU_panel_current) < cint(SECU_total) then transmit
			Loop until cint(SECU_panel_current) = cint(SECU_total)
		End if
		Next
	Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen SHEL_total, 1, 2, 78
		If SHEL_total <> 0 then
			member_number_designation = "Member " & HH_member & "- "
			row = 11
			Do
			EMReadScreen SHEL_amount, 8, row, 56
			If SHEL_amount <> "________" then
				EMReadScreen SHEL_type, 9, row, 24
				EMReadScreen SHEL_proof_check, 2, row, 67
				If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then
				SHEL_proof = ", no proof provided"
				Else
				SHEL_proof = ""
				End if
				SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ". ;"
			End if
			row = row + 1
			Loop until row = 19
			variable_written_to = variable_written_to & member_number_designation & SHEL_expense
		End if
		SHEL_expense = ""
		Next
	Elseif panel_read_from = "SWKR" then '---------------------------------------------------------------------------------------------------SWKR
		EMReadScreen SWKR_name, 35, 6, 32
		SWKR_name = replace(SWKR_name, "_", "")
		SWKR_name = split(SWKR_name)
		For each word in SWKR_name
		If word <> "" then
			first_letter_of_word = ucase(left(word, 1))
			rest_of_word = LCase(right(word, len(word) -1))
			If len(word) > 2 then
			variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
			Else
			variable_written_to = variable_written_to & word & " "
			End if
		End if
		Next
	Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen STWK_total, 1, 2, 78
		If STWK_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			EMReadScreen STWK_verification, 1, 7, 63
			If STWK_verification = "N" then
			STWK_verification = ", no proof provided"
			Else
			STWK_verification = ""
			End if
			EMReadScreen STWK_employer, 30, 6, 46
			STWK_employer = replace(STWK_employer, "_", "")
			STWK_employer = split(STWK_employer)
			For each STWK_part in STWK_employer
			If STWK_part <> "" then
				first_letter = ucase(left(STWK_part, 1))
				other_letters = LCase(right(STWK_part, len(STWK_part) -1))
				If len(STWK_part) > 3 then
				new_STWK_employer = new_STWK_employer & first_letter & other_letters & " "
				Else
				new_STWK_employer = new_STWK_employer & STWK_part & " "
				End if
			End if
			Next
			EMReadScreen STWK_income_stop_date, 8, 8, 46
			If STWK_income_stop_date = "__ __ __" then
			STWK_income_stop_date = "at unknown date"
			Else
			STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
			End if
		EMReadScreen voluntary_quit, 1, 10, 46
		vol_quit_info = ", Vol. Quit " & voluntary_quit
		IF voluntary_quit = "Y" THEN
			EMReadScreen good_cause, 1, 12, 67
			EMReadScreen fs_pwe, 1, 14, 46
			vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
		END IF
			variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
		End if
		new_STWK_employer = "" 'clearing variable to prevent duplicates
		Next
	Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		EMReadScreen UNEA_total, 1, 2, 78
		If UNEA_total <> 0 then
			variable_written_to = variable_written_to & "Member " & HH_member & "- "
			Do
			call HRF_add_UNEA_to_variable(variable_written_to)
			EMReadScreen UNEA_panel_current, 1, 2, 73
			If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
			Loop until cint(UNEA_panel_current) = cint(UNEA_total)
		End if
		Next
	Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
		For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		transmit
		EMReadScreen wreg_total, 1, 2, 78
		IF wreg_total <> "0" THEN
		EmWriteScreen "X", 13, 57
		transmit
		bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
		bene_yr_row = 10
		abawd_counted_months = 0
		second_abawd_period = 0
		month_count = 0
		DO
				'establishing variables for specific ABAWD counted month dates
				If bene_mo_col = "19" then counted_date_month = "01"
				If bene_mo_col = "23" then counted_date_month = "02"
				If bene_mo_col = "27" then counted_date_month = "03"
				If bene_mo_col = "31" then counted_date_month = "04"
				If bene_mo_col = "35" then counted_date_month = "05"
				If bene_mo_col = "39" then counted_date_month = "06"
				If bene_mo_col = "43" then counted_date_month = "07"
				If bene_mo_col = "47" then counted_date_month = "08"
				If bene_mo_col = "51" then counted_date_month = "09"
				If bene_mo_col = "55" then counted_date_month = "10"
				If bene_mo_col = "59" then counted_date_month = "11"
				If bene_mo_col = "63" then counted_date_month = "12"
				'reading to see if a month is counted month or not
				EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
				'counting and checking for counted ABAWD months
				IF is_counted_month = "X" or is_counted_month = "M" THEN
					EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
					abawd_counted_months_string = counted_date_month & "/" & counted_date_year
					abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
					abawd_counted_months = abawd_counted_months + 1				'adding counted months
				END IF

				'declaring & splitting the abawd months array
				If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
				abawd_months_array = Split(abawd_info_list, ",")

				'counting and checking for second set of ABAWD months
				IF is_counted_month = "Y" or is_counted_month = "N" THEN
					EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
					second_abawd_period = second_abawd_period + 1				'adding counted months
					second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
					second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
				END IF

				'declaring & splitting the second set of abawd months array
				If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
				second_months_array = Split(second_set_info_list,",")

				bene_mo_col = bene_mo_col - 4
				IF bene_mo_col = 15 THEN
					bene_yr_row = bene_yr_row - 1
					bene_mo_col = 63
				END IF
				month_count = month_count + 1
		LOOP until month_count = 36
			PF3

		EmreadScreen read_WREG_status, 2, 8, 50
		If read_WREG_status = "03" THEN  WREG_status = "WREG = incap"
		If read_WREG_status = "04" THEN  WREG_status = "WREG = resp for incap HH memb"
		If read_WREG_status = "05" THEN  WREG_status = "WREG = age 60+"
		If read_WREG_status = "06" THEN  WREG_status = "WREG = < age 16"
		If read_WREG_status = "07" THEN  WREG_status = "WREG = age 16-17, live w/prnt/crgvr"
		If read_WREG_status = "08" THEN  WREG_status = "WREG = resp for child < 6 yrs old"
		If read_WREG_status = "09" THEN  WREG_status = "WREG = empl 30 hrs/wk or equiv"
		If read_WREG_status = "10" THEN  WREG_status = "WREG = match grant part"
		If read_WREG_status = "11" THEN  WREG_status = "WREG = rec/app for unemp ins"
		If read_WREG_status = "12" THEN  WREG_status = "WREG = in schl, train prog or higher ed"
		If read_WREG_status = "13" THEN  WREG_status = "WREG = in CD prog"
		If read_WREG_status = "14" THEN  WREG_status = "WREG = rec MFIP"
		If read_WREG_status = "20" THEN  WREG_status = "WREG = pend/rec DWP or WB"
		If read_WREG_status = "22" THEN  WREG_status = "WREG = app for SSI"
		If read_WREG_status = "15" THEN  WREG_status = "WREG = age 16-17 not live w/ prnt/crgvr"
		If read_WREG_status = "16" THEN  WREG_status = "WREG = 50-59 yrs old"
		If read_WREG_status = "21" THEN  WREG_status = "WREG = resp for child < 18"
		If read_WREG_status = "17" THEN  WREG_status = "WREG = rec RCA or GA"
		If read_WREG_status = "18" THEN  WREG_status = "WREG = provide home schl"
		If read_WREG_status = "30" THEN  WREG_status = "WREG = mand FSET part"
		If read_WREG_status = "02" THEN  WREG_status = "WREG = non-coop w/ FSET"
		If read_WREG_status = "33" THEN  WREG_status = "WREG = non-coop w/ referral"
		If read_WREG_status = "__" THEN  WREG_status = "WREG = blank"

		EmreadScreen read_abawd_status, 2, 13, 50
		If read_abawd_status = "01" THEN  abawd_status = "ABAWD = work reg exempt."
			If read_abawd_status = "02" THEN  abawd_status = "ABAWD = < age 18."
		If read_abawd_status = "03" THEN  abawd_status = "ABAWD = age 50+."
		If read_abawd_status = "04" THEN  abawd_status = "ABAWD = crgvr of minor child."
		If read_abawd_status = "05" THEN  abawd_status = "ABAWD = pregnant."
		If read_abawd_status = "06" THEN  abawd_status = "ABAWD = emp ave 20 hrs/wk."
		If read_abawd_status = "07" THEN  abawd_status = "ABAWD = work exp participant."
		If read_abawd_status = "08" THEN  abawd_status = "ABAWD = othr E & T service."
		If read_abawd_status = "09" THEN  abawd_status = "ABAWD = reside in waiver area."
		IF read_abawd_status = "10" AND abawd_counted_months = "0" THEN
			abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo."
		Elseif read_abawd_status = "10" AND second_abawd_period = "0" THEN
			abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & "."
		Elseif read_abawd_status = "10" AND second_abawd_period <> "0" THEN
			abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo. Counted ABAWD months:" & abawd_info_list & ". Second set of ABAWD months used: " & second_abawd_period & ". Counted second set months: " & second_set_info_list & "."
		END IF
		If read_abawd_status = "11" THEN  abawd_status = "ABAWD = Using second set of ABAWD months. Counted second set months: " & second_set_info_list & "."
		If read_abawd_status = "12" THEN  abawd_status = "ABAWD = RCA or GA recip."
		'If read_abawd_status = "13" THEN  abawd_status = "ABAWD = ABAWD banked months."
		If read_abawd_status = "__" THEN  abawd_status = "ABAWD = blank"

		variable_written_to = variable_written_to & "Member " & HH_member & "- " & WREG_status & ", " & abawd_status & "; "
		END IF
		Next
	End if
	variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
	if right(variable_written_to, 1) = ";" then variable_written_to = left(variable_written_to, len(variable_written_to) - 1)
end function

function HRF_add_BUSI_to_variable(variable_name_for_BUSI)
	'--- This function adds STAT/BUSI data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
	'~~~~~ BUSI_variable: the variable used by the editbox you wish to autofill.
	'===== Keywords: MAXIS, autofill, BUSI
	'Error message handling
	BUSI_panel_error_message = ""

	'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
	EMReadScreen BUSI_footer_month, 5, 20, 55
	BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")

	'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
	If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		If BUSI_type = "10" then BUSI_type = "Lived Experience"
		EMWriteScreen "X", 7, 26
		EMSendKey "<enter>"
		EMWaitReady 0, 0
		If cash_check = 1 then
			EMReadScreen BUSI_ver, 1, 9, 73
		ElseIf HC_check = 1 then
			EMReadScreen BUSI_ver, 1, 12, 73
			If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
		ElseIf SNAP_check = 1 then
			EMReadScreen BUSI_ver, 1, 11, 73
		End if
		EMSendKey "<PF3>"
		EMWaitReady 0, 0
		If SNAP_check = 1 then
			EMReadScreen BUSI_amt, 8, 11, 68
			BUSI_amt = trim(BUSI_amt)
		ElseIf cash_check = 1 then
			EMReadScreen BUSI_amt, 8, 9, 54
			BUSI_amt = trim(BUSI_amt)
		ElseIf HC_check = 1 then
			EMWriteScreen "X", 17, 29
			EMSendKey "<enter>"
			EMWaitReady 0, 0
			EMReadScreen BUSI_amt, 8, 15, 54
			If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
			BUSI_amt = trim(BUSI_amt)
			EMSendKey "<PF3>"
			EMWaitReady 0, 0
		End if
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI"
		EMReadScreen BUSI_income_end_date, 8, 5, 71
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
		If IsDate(BUSI_income_end_date) = True then
			variable_name_for_BUSI = variable_name_for_BUSI & " (ended " & BUSI_income_end_date & ")"
		Else
			If BUSI_amt <> "" then variable_name_for_BUSI = variable_name_for_BUSI & ", ($" & BUSI_amt & "/monthly)"
		End if
		If BUSI_ver = "N" or BUSI_ver = "?" then
			variable_name_for_BUSI = variable_name_for_BUSI & ", no proof provided.; "
		Else
			variable_name_for_BUSI = variable_name_for_BUSI & ".; "
		End if
	Else		'------------This was updated 01/07/2015.
		'Checks the current footer month. If this is the future, it will know later on to read the HC pop-up
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		If datediff("d", date, BUSI_footer_month) > 0 then
			pull_future_HC = TRUE
		Else
			pull_future_HC = FALSE
		End if

		'Converting BUSI type code to a human-readable string
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		If BUSI_type = "10" then BUSI_type = "Lived Experience"

		'Reading and converting BUSI Self employment method into human-readable
		EMReadScreen BUSI_method, 2, 16, 53
		IF BUSI_method = "01" THEN BUSI_method = "50% Gross Income"
		IF BUSI_method = "02" THEN BUSI_method = "Tax Forms"

		'Going to the Gross Income Calculation pop-up
		EMWriteScreen "X", 6, 26
		transmit

		'Getting the verification codes for each type. Only does income, expenses are not included at this time.
		EMReadScreen BUSI_cash_ver, 1, 9, 73
		EMReadScreen BUSI_IVE_ver, 1, 10, 73
		EMReadScreen BUSI_SNAP_ver, 1, 11, 73
		EMReadScreen BUSI_HCA_ver, 1, 12, 73
		EMReadScreen BUSI_HCB_ver, 1, 13, 73

		'Converts each ver type to human readable
		If BUSI_cash_ver = "1" then BUSI_cash_ver = "tax returns provided"
		If BUSI_cash_ver = "2" then BUSI_cash_ver = "receipts provided"
		If BUSI_cash_ver = "3" then BUSI_cash_ver = "client ledger provided"
		If BUSI_cash_ver = "6" then BUSI_cash_ver = "other doc provided"
		If BUSI_cash_ver = "N" then BUSI_cash_ver = "no proof provided"
		If BUSI_cash_ver = "?" then BUSI_cash_ver = "no proof provided"
		If BUSI_IVE_ver = "1" then BUSI_IVE_ver = "tax returns provided"
		If BUSI_IVE_ver = "2" then BUSI_IVE_ver = "receipts provided"
		If BUSI_IVE_ver = "3" then BUSI_IVE_ver = "client ledger provided"
		If BUSI_IVE_ver = "6" then BUSI_IVE_ver = "other doc provided"
		If BUSI_IVE_ver = "N" then BUSI_IVE_ver = "no proof provided"
		If BUSI_IVE_ver = "?" then BUSI_IVE_ver = "no proof provided"
		If BUSI_SNAP_ver = "1" then BUSI_SNAP_ver = "tax returns provided"
		If BUSI_SNAP_ver = "2" then BUSI_SNAP_ver = "receipts provided"
		If BUSI_SNAP_ver = "3" then BUSI_SNAP_ver = "client ledger provided"
		If BUSI_SNAP_ver = "6" then BUSI_SNAP_ver = "other doc provided"
		If BUSI_SNAP_ver = "N" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_SNAP_ver = "?" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_HCA_ver = "1" then BUSI_HCA_ver = "tax returns provided"
		If BUSI_HCA_ver = "2" then BUSI_HCA_ver = "receipts provided"
		If BUSI_HCA_ver = "3" then BUSI_HCA_ver = "client ledger provided"
		If BUSI_HCA_ver = "6" then BUSI_HCA_ver = "other doc provided"
		If BUSI_HCA_ver = "N" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCA_ver = "?" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCB_ver = "1" then BUSI_HCB_ver = "tax returns provided"
		If BUSI_HCB_ver = "2" then BUSI_HCB_ver = "receipts provided"
		If BUSI_HCB_ver = "3" then BUSI_HCB_ver = "client ledger provided"
		If BUSI_HCB_ver = "6" then BUSI_HCB_ver = "other doc provided"
		If BUSI_HCB_ver = "N" then BUSI_HCB_ver = "no proof provided"
		If BUSI_HCB_ver = "?" then BUSI_HCB_ver = "no proof provided"

		'Back to the main screen
		PF3

		'Reading each income amount, trimming them to clean out unneeded spaces.
		EMReadScreen BUSI_cash_retro_amt, 8, 8, 55
		BUSI_cash_retro_amt = trim(BUSI_cash_retro_amt)
		EMReadScreen BUSI_cash_pro_amt, 8, 8, 69
		BUSI_cash_pro_amt = trim(BUSI_cash_pro_amt)
		EMReadScreen BUSI_IVE_amt, 8, 9, 69
		BUSI_IVE_amt = trim(BUSI_IVE_amt)
		EMReadScreen BUSI_SNAP_retro_amt, 8, 10, 55
		BUSI_SNAP_retro_amt = trim(BUSI_SNAP_retro_amt)
		EMReadScreen BUSI_SNAP_pro_amt, 8, 10, 69
		BUSI_SNAP_pro_amt = trim(BUSI_SNAP_pro_amt)

		'Handling to ensure that Retro and Prosp match, error handling if not
		If BUSI_SNAP_retro_amt <> BUSI_SNAP_pro_amt Then BUSI_panel_error_message = BUSI_panel_error_message & "The amounts entered in the retrospective and prospective fields do not match. They should all be the same. Please update and then rerun this script."

		'Pulls prospective amounts for HC, either from prosp side or from HC inc est.
		If pull_future_HC = False then
			EMReadScreen BUSI_HCA_amt, 8, 11, 69
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 12, 69
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
		Else
			EMWriteScreen "x", 17, 27
			transmit
			EMReadScreen BUSI_HCA_amt, 8, 15, 54
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 16, 54
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
			PF3
		End if

		'Reads end date logic (in case it ended), converts to an actual date
		EMReadScreen BUSI_income_end_date, 8, 5, 72
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")

		'Entering the variable details based on above
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI:; "
		If IsDate(BUSI_income_end_date) = True then	variable_name_for_BUSI = variable_name_for_BUSI & "- Income ended " & BUSI_income_end_date & ".; "
		If BUSI_cash_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH retro: $" & BUSI_cash_retro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_cash_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH pro: $" & BUSI_cash_pro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_IVE_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- IV-E: $" & BUSI_IVE_amt & " budgeted, " & BUSI_IVE_ver & "; "
		If BUSI_SNAP_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP retro: $" & BUSI_SNAP_retro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_SNAP_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP pro: $" & BUSI_SNAP_pro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_HCA_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method A: $" & BUSI_HCA_amt & " budgeted, " & BUSI_HCA_ver & "; "
		If BUSI_HCB_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method B: $" & BUSI_HCB_amt & " budgeted, " & BUSI_HCB_ver & "; "
		'Checks to see if pre 01/15 or post 02/15 then decides what to put in case note based on what was found/needed on the self employment method.
		If IsDate(BUSI_income_end_date) = false then
			IF BUSI_method <> "__" or BUSI_method = "" THEN
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: " & BUSI_method & "; "
			Else
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: None; "
			END IF
		End if
	End if
	If BUSI_panel_error_message <> "" Then script_end_procedure(BUSI_panel_error_message)
end function

function HRF_add_JOBS_to_variable(variable_name_for_JOBS)
	'--- This function adds STAT/JOBS data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
	'~~~~~ JOBS_variable: the variable used by the editbox you wish to autofill.
	'===== Keywords: MAXIS, autofill, JOBS
	'Set error message handling 
	JOBS_panel_error_message = ""

	EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
	JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
	EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
	JOBS_type = replace(JOBS_type, "_", ""	)
	JOBS_type = trim(JOBS_type)
	JOBS_type = split(JOBS_type)
	For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
		If JOBS_part <> "" then
		first_letter = ucase(left(JOBS_part, 1))
		other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
		new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
		End if
	Next
	'Read if pay frequency set to 1 on JOBS panel
	EMReadScreen JOBS_pay_frequency, 1, 18, 35
	If JOBS_pay_frequency <> "1" then JOBS_panel_error_message = JOBS_panel_error_message & "The pay frequency must be changed to 1 on the JOBS panel. It is currently " & JOBS_pay_frequency & "." & VbCR & vbCr
	
	'Read the prospective and retrospective amounts. Ensure they are the same and that there are no lines filled out besides the first
	EmReadScreen JOBS_panel_retro_pay_amount, 8, 12, 38
	JOBS_panel_retro_pay_amount = trim(JOBS_panel_retro_pay_amount)
	EmReadScreen JOBS_panel_retro_pay_amount_line_2_check, 8, 13, 38
	EmReadScreen JOBS_panel_prosp_pay_amount, 8, 12, 67
	JOBS_panel_prosp_pay_amount = trim(JOBS_panel_prosp_pay_amount)
	EmReadScreen JOBS_panel_prosp_pay_amount_line_2_check, 8, 13, 67
	
	If JOBS_panel_retro_pay_amount_line_2_check <> "________" or JOBS_panel_prosp_pay_amount_line_2_check <> "________" Then JOBS_panel_error_message = JOBS_panel_error_message & "There is pay information on the second line of the retrospective and/or prospective fields on the JOBS panel. There should only be pay information on the first line. Please update and then rerun this script." & VbCR & vbCr

	' EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
	' jobs_hourly_wage = replace(jobs_hourly_wage, "_", "")   'trimming any underscores
	' Navigates to the FS PIC
	EMWriteScreen "X", 19, 38
	transmit
	' EMReadScreen SNAP_JOBS_amt, 8, 17, 56
	' SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
	EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
	jobs_SNAP_prospective_amt = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
	' EMReadScreen snap_pay_frequency, 1, 5, 64
	' EMReadScreen date_of_pic_calc, 8, 5, 34
	' date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
	transmit

	If JOBS_panel_retro_pay_amount <> JOBS_panel_prosp_pay_amount or JOBS_panel_retro_pay_amount <> jobs_SNAP_prospective_amt or JOBS_panel_prosp_pay_amount <> jobs_SNAP_prospective_amt Then UNEA_panel_error_message = UNEA_panel_error_message & "The amount entered in the PIC for the monthly prospective income and the prospective and retrospective pay amounts entered on the panel are not all the same amount. They should all be the same. Please update and then rerun this script." & VbCR & vbCr

	'Navigates to GRH PIC
	EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
	IF GRH_PIC_check = "GRH" THEN
		EMWriteScreen "X", 19, 71
		transmit
		EMReadScreen GRH_JOBS_amt, 8, 16, 69
		GRH_JOBS_amt = trim(GRH_JOBS_amt)
		EMReadScreen GRH_pay_frequency, 1, 3, 63
		EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
		GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
		PF3
	END IF
	'  Reads the information on the retro side of JOBS
	EMReadScreen retro_JOBS_amt, 8, 17, 38
	retro_JOBS_amt = trim(retro_JOBS_amt)
	'  Reads the information on the prospective side of JOBS
	EMReadScreen prospective_JOBS_amt, 8, 17, 67
	prospective_JOBS_amt = trim(prospective_JOBS_amt)
	'  Reads the information about health care off of HC Income Estimator
	EMReadScreen pay_frequency, 1, 18, 35
	EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
	IF HC_income_est_check = "Est" Then 'this is the old position
		EMWriteScreen "X", 19, 54
	ELSE								'this is the new position
		EMWriteScreen "X", 19, 48
	END IF
	transmit
	EMReadScreen HC_JOBS_amt, 8, 11, 63
	HC_JOBS_amt = trim(HC_JOBS_amt)
	transmit

	EMReadScreen JOBS_ver, 1, 6, 38
	EMReadScreen JOBS_income_end_date, 8, 9, 49
		'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
	If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
	If IsDate(JOBS_income_end_date) = True then
		variable_name_for_JOBS = variable_name_for_JOBS & new_JOBS_type & "(ended " & JOBS_income_end_date & "); "
	Else
		If pay_frequency = "1" then pay_frequency = "monthly"
		If pay_frequency = "2" then pay_frequency = "semimonthly"
		If pay_frequency = "3" then pay_frequency = "biweekly"
		If pay_frequency = "4" then pay_frequency = "weekly"
		If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
		If pay_frequency <> "monthly" then JOBS_panel_error_message = JOBS_panel_error_message & "The pay frequency is not currently monthly (1). Update the panel to reflect a monthly (1) frequency. Please update and then rerun this script." & VbCR & vbCr
		IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
		IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
		IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
		IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
		IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
		If GRH_pay_frequency = "1" then GRH_pay_frequency = "monthly"
		If GRH_pay_frequency = "2" then GRH_pay_frequency = "semimonthly"
		If GRH_pay_frequency = "3" then GRH_pay_frequency = "biweekly"
		If GRH_pay_frequency = "4" then GRH_pay_frequency = "weekly"
		variable_name_for_JOBS = variable_name_for_JOBS & "EI from " & trim(new_JOBS_type) & ", " & JOBS_month  & " amt: "
		' If SNAP_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- SNAP PIC: $" & SNAP_JOBS_amt & "/" & snap_pay_frequency & ", SNAP PIC Prospective: $" & jobs_SNAP_prospective_amt & ", calculated " & date_of_pic_calc & "; "
		IF JOBS_panel_prosp_pay_amount <> "" THEN variable_name_for_JOBS = variable_name_for_JOBS & "Prosp: $" & JOBS_panel_prosp_pay_amount & " total; "
		If GRH_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- GRH PIC: $" & GRH_JOBS_amt & "/" & GRH_pay_frequency & ", calculated " & GRH_date_of_pic_calc & "; "
		' If retro_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- Retrospective: $" & retro_JOBS_amt & " total; "
		' IF isnumeric(jobs_hourly_wage) THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Hourly Wage: $" & jobs_hourly_wage & "; "
		'Leaving out HC income estimator if footer month is not Current month + 1
		current_month_for_hc_est = dateadd("m", "1", date)
		current_month_for_hc_est = datepart("m", current_month_for_hc_est)
		IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
		IF MAXIS_footer_month = current_month_for_hc_est THEN
			IF HC_JOBS_amt <> "________" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- HC Inc Est: $" & HC_JOBS_amt & "/" & pay_frequency & "; "
		End If
		If JOBS_ver = "N" or JOBS_ver = "?" then variable_name_for_JOBS = variable_name_for_JOBS & "- No proof provided for this panel; "
	End if
	If JOBS_panel_error_message <> "" Then script_end_procedure(JOBS_panel_error_message)
end function

function HRF_add_UNEA_to_variable(variable_name_for_UNEA)
	'--- This function adds STAT/UNEA data to a variable, which can then be displayed in a dialog. See autofill_editbox_from_MAXIS.
	'~~~~~ UNEA_variable: the variable used by the editbox you wish to autofill.
	'===== Keywords: MAXIS, autofill, UNEA
	'Error handling message
	UNEA_panel_error_message = ""

	EMReadScreen UNEA_month, 5, 20, 55
	UNEA_month = replace(UNEA_month, " ", "/")
	EMReadScreen UNEA_type, 16, 5, 40
	If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
	If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
	If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
	UNEA_type = trim(UNEA_type)
	EMReadScreen UNEA_ver, 1, 5, 65
	EMReadScreen UNEA_income_end_date, 8, 7, 68
	If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
	If IsDate(UNEA_income_end_date) = True then
		variable_name_for_UNEA = variable_name_for_UNEA & UNEA_type & " (ended " & UNEA_income_end_date & "); "
	Else
		' EMReadScreen UNEA_amt, 8, 18, 68
		' UNEA_amt = trim(UNEA_amt)
		EMWriteScreen "X", 10, 26
		transmit
		' EMReadScreen SNAP_UNEA_amt, 8, 17, 56
		' SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
		' EMReadScreen snap_pay_frequency, 1, 5, 64
		' EMReadScreen date_of_pic_calc, 8, 5, 34
		' date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
		
		EMReadScreen UNEA_PIC_prosp_amt, 8, 18, 56
		UNEA_PIC_prosp_amt = trim(UNEA_PIC_prosp_amt)
		transmit

		' EMReadScreen retro_UNEA_amt, 8, 18, 39
		' retro_UNEA_amt = trim(retro_UNEA_amt)
		' EMReadScreen prosp_UNEA_amt, 8, 18, 68
		' prosp_UNEA_amt = trim(prosp_UNEA_amt)
		EMWriteScreen "X", 6, 56
		transmit
		EMReadScreen HC_UNEA_amt, 8, 9, 65
		HC_UNEA_amt = trim(HC_UNEA_amt)
		EMReadScreen pay_frequency, 1, 10, 63
		transmit

		'Read the prospective and retrospective amounts. Ensure they are the same and that there are no lines filled out besides the first
		EmReadScreen UNEA_panel_retro_pay_amount, 8, 13, 39
		UNEA_panel_retro_pay_amount = trim(UNEA_panel_retro_pay_amount)
		EmReadScreen UNEA_panel_retro_pay_amount_line_2_check, 8, 14, 39
		EmReadScreen UNEA_panel_prosp_pay_amount, 8, 13, 68
		UNEA_panel_prosp_pay_amount = trim(UNEA_panel_prosp_pay_amount)
		EmReadScreen UNEA_panel_prosp_pay_amount_line_2_check, 8, 14, 68

		If UNEA_panel_retro_pay_amount <> UNEA_panel_prosp_pay_amount or UNEA_panel_retro_pay_amount <> UNEA_PIC_prosp_amt or UNEA_panel_prosp_pay_amount <> UNEA_PIC_prosp_amt Then UNEA_panel_error_message = UNEA_panel_error_message & "The amount entered in the PIC for the monthly prospective income and the prospective and retrospective pay amounts entered on the panel are not all the same amount. They should all be the same. Please update and then rerun this script." & VbCR & vbCr

		If UNEA_panel_retro_pay_amount_line_2_check <> "________" or UNEA_panel_prosp_pay_amount_line_2_check <> "________" Then UNEA_panel_error_message = UNEA_panel_error_message & "There is pay information on the second line of the retrospective and/or prospective fields on the UNEA panel. There should only be pay information on the first line. Please update and then rerun this script." & VbCR & vbCr

		If HC_UNEA_amt = "________" then
			EMReadScreen HC_UNEA_amt, 8, 18, 68
			HC_UNEA_amt = trim(HC_UNEA_amt)
			pay_frequency = "mo budgeted prospectively"
		End If
		If pay_frequency = "1" then pay_frequency = "monthly"
		If pay_frequency = "2" then pay_frequency = "semimonthly"
		If pay_frequency = "3" then pay_frequency = "biweekly"
		If pay_frequency = "4" then pay_frequency = "weekly"
		If pay_frequency = "_" then pay_frequency = "non-monthly"
		IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
		IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
		IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
		IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
		IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
		variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amt: "
		' If SNAP_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
		' If retro_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
		If UNEA_panel_prosp_pay_amount <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "Prosp: $" & UNEA_panel_prosp_pay_amount & " total; "
		'Leaving out HC income estimator if footer month is not Current month + 1
		current_month_for_hc_est = dateadd("m", "1", date)
		current_month_for_hc_est = datepart("m", current_month_for_hc_est)
		IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
		IF MAXIS_footer_month = current_month_for_hc_est THEN
			If HC_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- HC Inc Est: $" & HC_UNEA_amt & "/" & pay_frequency & "; "
		END IF
		If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "
	End if
	If UNEA_panel_error_message <> "" Then script_end_procedure(UNEA_panel_error_message)
end function

Call MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing case number & footer month/year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'New dialog alerting worker to changes with implementation of six-month reporting for MFIP and GA
Dialog1 = "" 'blanking out dialog name
BeginDialog Dialog1, 0, 0, 191, 75, "HRF - Notice of Changes"
  ButtonGroup ButtonPressed
    OkButton 85, 55, 50, 15
    CancelButton 135, 55, 50, 15
    PushButton 5, 55, 65, 15, "Processing Guide", guide_btn						
  Text 5, 5, 35, 10, "NOTICE:"
  Text 5, 20, 180, 30, "Six-month budgeting will begin for MFIP, UHFS, and GA effective benefit month 03/25. Please review the processing guide linked below as needed."
EndDialog

DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation
		If ButtonPressed = guide_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhssir.cty.dhs.state.mn.us/Shared%20Documents/Guide%20to%20Six-Month%20Budgeting%20Final%20Version.pdf"	'copy the instructions URL here
			err_msg = "LOOP"
		End If
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

Do
    Do
        '-------------------------------------------------------------------------------------------------DIALOG
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 181, 110, "HRF Case Number"
          EditBox 80, 5, 70, 15, MAXIS_case_number
          EditBox 80, 25, 18, 15, MAXIS_footer_month
          EditBox 106, 25, 18, 15, MAXIS_footer_year
          CheckBox 10, 70, 30, 10, "MFIP", MFIP_check
          CheckBox 45, 70, 30, 10, "SNAP", SNAP_check
          CheckBox 85, 70, 20, 10, "HC", HC_check
          CheckBox 115, 70, 25, 10, "GA", GA_check
          CheckBox 145, 70, 50, 10, "MSA", MSA_check
          ButtonGroup ButtonPressed
            OkButton 35, 90, 50, 15
            CancelButton 95, 90, 50, 15
          Text 30, 10, 50, 10, "Case number:"
          Text 5, 30, 75, 10, "Footer month (MM/YY):"
		  Text 101, 30, 4, 10, "/"
          Text 80, 40, 75, 10, "(benefit month)"
          GroupBox 5, 55, 170, 30, "Programs Recertifying"
        EndDialog

        err_msg = ""
      	Dialog Dialog1
      	cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        If (MFIP_check = 0 and SNAP_check = 0 and HC_check = 0 and GA_check = 0 and MSA_check = 0) then err_msg = err_msg & "* Select all applicable programs at monthly report."

		'Checking for PRIV cases.
		EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
		IF priv_check = "PRIVIL" THEN script_end_procedure("This case is a privliged case. You do not have access to this case.")

        'Checking to ensure the case is actually at a HRF
        Call check_for_MAXIS(False)
        Call navigate_to_MAXIS_screen("STAT", "MONT")
        EmReadscreen HRF_panels, 1, 2, 73
        If HRF_panels = 0 then
            script_end_procedure_with_error_report("This case is not subject to monthly reporting. The script will now end.")
        Else
            cash_hrf = False    'defaulting programs to false for determination
            hc_hrf = False
            snap_hrf = False

            'setting boolean here since there are more than one cash programs
            If MFIP_check = 1 or GA_check = 1 or MSA_check = 1 then
                cash_progs = True
            else
                cash_progs = False
            End if

            EmReadscreen cash_code, 1, 11, 43
            EmReadscreen snap_code, 1, 11, 53
            EmReadscreen HC_code, 1, 11, 63

            If cash_code <> "_" then cash_HRF = True
            If snap_code <> "_" then snap_HRF = True
            If HC_code <> "_" then hc_HRF = True

            'program selected in dialog, not open available as HRF process
            If HC_check = 1 and hc_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process health care, but you cannot use a HRF for health care programs on this case. Update your program selections." & vbcr
            If SNAP_check = 1 and snap_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process SNAP, but you cannot use a HRF for SNAP on this case. Update your program selections." & vbcr
            If cash_progs = True and cash_hrf = False then err_msg = err_msg & vbcr & "* You've selected to process cash programs, but you cannot use a HRF for cash programs on this case. Update your program selection." & vbcr

            'program listed on MONT page, but NOT in program selection in dialog
            If HC_check = 0 and hc_hrf = True then err_msg = err_msg & vbcr & "* Health Care is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
            If SNAP_check = 0 and snap_hrf = True then err_msg = err_msg & vbcr & "* SNAP is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
            If cash_progs = false and cash_hrf = True then err_msg = err_msg & vbcr & "* Cash is not selected on this case, but is due for a HRF this month. Update your program selections." & vbcr
        End if
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Convert the benefit footer month to date
benefit_footer_month_year_date = dateadd("d", 0, MAXIS_footer_month & "/01/20" & MAXIS_footer_year)

'Create list of selected programs
programs_selected_list = "*"
If MFIP_check = 1 then programs_selected_list = programs_selected_list & "MFIP*"
If SNAP_check = 1 then programs_selected_list = programs_selected_list & "SNAP*"
If SNAP_check = 1 then programs_selected_list = programs_selected_list & "HC*"
If GA_check = 1 then programs_selected_list = programs_selected_list & "GA*"
If MSA_check = 1 then programs_selected_list = programs_selected_list & "MSA*"

If datediff("d", benefit_footer_month_year_date, #03/01/2025#) > 0 or (instr(programs_selected_list, "MFIP") = 0 and instr(programs_selected_list, "SNAP") = 0 and instr(programs_selected_list, "GA") = 0) Then
	'Benefit month BEFORE 03 25 OR programs selected are HC or MSA only (no MFIP, SNAP, OR GA)

	'NAV to STAT
	call navigate_to_MAXIS_screen("STAT", "MEMB")

	'Creating a custom dialog for determining who the HH members are
	call HH_member_custom_dialog(HH_member_array)

	'Autofilling info for case note
	call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
	call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
	call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
	call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
	call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
	call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

	'Cleaning up info for case note
	HRF_computer_friendly_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
	retro_month_name = monthname(datepart("m", (dateadd("m", -2, HRF_computer_friendly_month))))
	pro_month_name = monthname(datepart("m", (HRF_computer_friendly_month)))
	HRF_month = retro_month_name & "/" & pro_month_name
	next_month_hrf_not_received_checkbox = unchecked
	next_retro_month_name = monthname(datepart("m", (dateadd("m", -1, HRF_computer_friendly_month))))
	next_month_name = monthname(datepart("m", DateAdd("m", 1, HRF_computer_friendly_month)))
	next_HRF_month = next_retro_month_name & "/" & next_month_name

	'If a HRF is being run for a HC case, script will ask if this is a LTC case
	If HC_check = checked Then
		'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
		LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have different fields in their dialog.", vbYesNoCancel)
		If LTC_case = vbCancel then stopscript
	Else
		LTC_case = vbNo
	End If

	'If workers answers yes to this is a LTC case - script runs this specific functionality
	If LTC_case = vbYes then
		'LTC cases should not have these programs active
		If MFIP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* MFIP will be removed."
		If SNAP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* SNAP will be removed."
		If GA_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* GA will be removed."

		'Alerting worker that these programs will be unchecked.
		If uncheck_msg <> "" Then MsgBox "You have checked programs that should not be active with LTC. These programs will not be added to the note." & vbNewLine & uncheck_msg

		MFIP_check = unchecked
		SNAP_check = unchecked
		GA_check = unchecked

		'Getting some additional information for the dialog to be autofilled
		call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
		call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
		call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)

		'Going to find the current facility to autofil the dialog
		Call navigate_to_MAXIS_screen ("STAT", "FACI")

		'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
		Do
			EMReadScreen FACI_current_panel, 1, 2, 73
			EMReadScreen FACI_total_check, 1, 2, 78
			EMReadScreen in_year_check_01, 4, 14, 53
			EMReadScreen in_year_check_02, 4, 15, 53
			EMReadScreen in_year_check_03, 4, 16, 53
			EMReadScreen in_year_check_04, 4, 17, 53
			EMReadScreen in_year_check_05, 4, 18, 53
			EMReadScreen out_year_check_01, 4, 14, 77
			EMReadScreen out_year_check_02, 4, 15, 77
			EMReadScreen out_year_check_03, 4, 16, 77
			EMReadScreen out_year_check_04, 4, 17, 77
			EMReadScreen out_year_check_05, 4, 18, 77
			If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or (in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
				currently_in_FACI = True
				exit do
			Elseif FACI_current_panel = FACI_total_check then
				currently_in_FACI = False
				exit do
			Else
				transmit
			End if
		Loop until FACI_current_panel = FACI_total_check

		If in_year_check_01 <> "____" and out_year_check_01 = "____" Then EMReadScreen date_in, 10, 14, 47
		If in_year_check_02 <> "____" and out_year_check_02 = "____" Then EMReadScreen date_in, 10, 15, 47
		If in_year_check_03 <> "____" and out_year_check_03 = "____" Then EMReadScreen date_in, 10, 16, 47
		If in_year_check_04 <> "____" and out_year_check_04 = "____" Then EMReadScreen date_in, 10, 17, 47
		If in_year_check_05 <> "____" and out_year_check_05 = "____" Then EMReadScreen date_in, 10, 18, 47

		admit_date = replace(date_in, " ", "/")

		'Gets Facility name and admit in date and enters it into the dialog
		If currently_in_FACI = True then
			EMReadScreen FACI_name, 30, 6, 43
			facility_info = trim(replace(FACI_name, "_", ""))
		End if

		'confirms that case is in the footer month/year selected by the user
		Call MAXIS_footer_month_confirmation
		Call MAXIS_background_check

		'Goes to STAT WKEX to get deductions and possible FIAT reasons to autofil the dialog
		Call navigate_to_MAXIS_screen("STAT", "WKEX")
		EMReadScreen WKEX_check, 1, 2, 73
		If WKEX_check = "0" then
			script_end_procedure("You do not have a WKEX panel. Please create a WKEX panel for your HC case, and re-run the script.")
		Elseif WKEX_check <> "0" then
			'Reads work expenses for MEMB 01, verification codes and impairment related code
			EMReadScreen program_check,          2, 5, 33
			EMReadScreen federal_tax,            8, 7, 57
			EMReadScreen federal_tax_verif_code, 1, 7, 69
			EMReadScreen state_tax,              8, 8, 57
			EMReadScreen state_tax_verif_code, 1, 8, 69
			EMReadScreen FICA_witheld,           8, 9, 57
			EMReadScreen FICA_witheld_verif_code, 1, 9, 69
			EMReadScreen transportation_expense, 8, 10, 57
			EMReadScreen transportation_expense_verif_code, 1, 10, 69
			EMReadScreen transportation_impair, 1, 10, 75
			EMReadScreen meals_expense, 8, 11, 57
			EMReadScreen meals_impair, 1, 11, 75
			EMReadScreen meals_expense_verif_code, 1, 11, 69
			EMReadScreen uniform_expense, 8, 12, 57
			EMReadScreen uniform_expense_verif_code, 1, 12, 69
			EMReadScreen uniform_impair, 1, 12, 75
			EMReadScreen tools_expense, 8, 13, 57
			EMReadScreen tools_expense_verif_code, 1, 13, 69
			EMReadScreen tools_impair, 1, 13, 75
			EMReadScreen dues_expense, 8, 14, 57
			EMReadScreen dues_expense_verif_code, 1, 14, 69
			EMReadScreen dues_impair, 1, 14, 75
			EMReadScreen other_expense, 8, 15,	57
			EMReadScreen other_expense_verif_code, 1, 15, 69
			EMReadScreen other_impair, 1, 15, 75
		End IF

		'cleaning up the WKEX variables
		federal_tax = replace(federal_tax, "_", "")
		federal_tax = trim(federal_tax)
		state_tax = replace(state_tax, "_", "")
		state_tax = trim(state_tax)
		FICA_witheld = replace(FICA_witheld, "_", "")
		FICA_witheld = trim(FICA_witheld)
		transportation_expense = replace(transportation_expense, "_", "")
		transportation_expense = trim(transportation_expense)
		meals_expense = replace(meals_expense, "_", "")
		meals_expense = trim(meals_expense)
		uniform_expense = replace(uniform_expense, "_", "")
		uniform_expense = trim(uniform_expense)
		tools_expense = replace(tools_expense, "_", "")
		tools_expense = trim(tools_expense)
		dues_expense = replace(dues_expense, "_", "")
		dues_expense = trim(dues_expense)
		other_expense = replace(other_expense, "_", "")
		other_expense = trim(other_expense)

		'Gives unverified expenses and blank expenses the value of $0 and adds non-zero amounts to the dialog for autofil
		If federal_tax = "" OR federal_tax_verif_code = "N" then
			federal_tax = "0"
		Else
			hc_deductions = hc_deductions & "; Federal Tax - $" & federal_tax
		End if
		If state_tax = "" OR state_tax_verif_code = "N" then
			state_tax = "0"
		Else
			hc_deductions = hc_deductions & "; State Tax - $" & state_tax
		End if
		If FICA_witheld = "" OR FICA_witheld_verif_code = "N" then
			FICA_witheld = "0"
		Else
			hc_deductions = hc_deductions & "; FICA - $" & FICA_witheld
		End if
		If transportation_expense = "" OR transportation_expense_verif_code = "N" OR transportation_impair =  "_" OR transportation_impair = "N" then
			transportation_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Transportation Expense - $" & transportation_expense
		End if
		If meals_expense = "" OR meals_expense_verif_code = "N" OR meals_impair = "_" OR meals_impair = "N" then
			meals_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Meals Expense - $" & meals_expense
		End if
		If uniform_expense = "" OR uniform_expense_verif_code = "N" OR uniform_impair = "_" OR uniform_impair = "N" then
			uniform_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Uniform Expense - $" & uniform_expense
		End if
		If tools_expense = "" OR tools_expense_verif_code = "N" OR tools_impair = "_" OR tools_impair = "N" then
			tools_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Tools Expense - $" & tools_expense
		End if
		If dues_expense = "" OR dues_expense_verif_code = "N" OR dues_impair = "_" OR dues_impair = "N" then
			dues_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Dues Expense - $" & dues_expense
		End if
		If other_expense = "" OR other_expense_verif_code = "N" OR other_impair = "_" OR other_impair = "N" then
			other_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Other Expense - $" & other_expense
		End if

		'Checks PDED other expenses, will need to add PDED and WKEX other expenses together
		Call navigate_to_MAXIS_screen("STAT", "PDED")
		EMReadScreen other_earned_income_PDED, 8, 11, 62

		'cleaning up PDED variables
		other_earned_income_PDED = replace(other_earned_income_PDED, "_", "")
		other_earned_income_PDED = trim(other_earned_income_PDED)

		'Gives blank expenses the value of $0
		If other_earned_income_PDED = "" then
			other_earned_income_PDED = "0"
		Else
			hc_deductions = hc_deductions & "; Other Earned Income Deductions - $" & other_earned_income_PDED
		End If

		'Determining if earned income is less than $80
		Call navigate_to_MAXIS_screen ("STAT", "JOBS")
		EMReadScreen JOBS_panel_income, 7, 17, 68
		JOBS_panel_income = trim(JOBS_panel_income)
		If IsNumeric(JOBS_panel_income) = TRUE Then
			If abs(JOBS_panel_income) < 80 then
				special_pers_allow = JOBS_panel_income	'if less then $80 deduction is earned income amount
			ELSE
				special_pers_allow = "80.00"		'otherwise deduction is $80
			END IF
		Else
			JOBS_panel_income = ""
		End If

		If JOBS_panel_income <> "" Then hc_deductions = hc_deductions & "; Special Allowance - $" & special_pers_allow

		'All of the deductions found to this point need to be FIATed. Added these to the FIAT varirable.
		FIAT_reasons = hc_deductions

		'Going to see if there is a deduction on MEDI. (This does not have to be FIATED)
		Call navigate_to_MAXIS_screen ("STAT", "MEDI")
		EMReadScreen medi_panel_exists, 1, 2, 78
		If medi_panel_exists = "1" Then
			EMReadScreen part_b_premium, 9, 7, 72
			part_b_premium = trim(part_b_premium)
			If part_b_premium <> "________" Then hc_deductions = hc_deductions & "; Medicare Premium - $" & part_b_premium
		End If

		'Formatting the variables for the dialog
		hc_deductions = right(hc_deductions, len(hc_deductions) - 2)
		FIAT_reasons = right(FIAT_reasons, len(FIAT_reasons) - 2)

		'The case note dialog, complete with panel navigation, reading the ELIG/MSA or ELIG/HC screen, and navigation to case note, as well as logic for certain sections to be required.
		DO
			DO
				Do
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					BeginDialog Dialog1, 0, 0, 451, 295, "HRF for LTC Cases"
					EditBox 65, 10, 85, 15, HRF_datestamp
					DropListBox 240, 10, 80, 15, "complete"+chr(9)+"incomplete", HRF_status
					EditBox 50, 30, 165, 15, facility_info
					EditBox 280, 30, 55, 15, admit_date
					CheckBox 350, 5, 80, 10, "Sent 3050 to Facility", sent_3050_checkbox
					CheckBox 350, 20, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
					CheckBox 350, 35, 80, 10, "Next HRF Released", HRF_release_checkbox
					EditBox 65, 50, 380, 15, earned_income
					EditBox 70, 70, 375, 15, unearned_income
					EditBox 110, 90, 335, 15, notes_on_income
					EditBox 40, 110, 405, 15, assets
					EditBox 50, 130, 395, 15, hc_deductions
					EditBox 100, 150, 345, 15, FIAT_reasons
					EditBox 50, 170, 395, 15, other_notes
					EditBox 235, 190, 210, 15, verifs_needed
					EditBox 235, 210, 210, 15, actions_taken
					EditBox 165, 275, 105, 15, worker_signature
					ButtonGroup ButtonPressed
						OkButton 340, 275, 50, 15
						CancelButton 390, 275, 50, 15
						PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
						PushButton 10, 205, 25, 10, "BUSI", BUSI_button
						PushButton 35, 205, 25, 10, "JOBS", JOBS_button
						PushButton 10, 215, 25, 10, "RBIC", RBIC_button
						PushButton 35, 215, 25, 10, "UNEA", UNEA_button
						PushButton 75, 205, 25, 10, "ACCT", ACCT_button
						PushButton 100, 205, 25, 10, "CARS", CARS_button
						PushButton 125, 205, 25, 10, "CASH", CASH_button
						PushButton 150, 205, 25, 10, "OTHR", OTHR_button
						PushButton 75, 215, 25, 10, "REST", REST_button
						PushButton 100, 215, 25, 10, "SECU", SECU_button
						PushButton 125, 215, 25, 10, "TRAN", TRAN_button
						PushButton 10, 250, 25, 10, "MEMB", MEMB_button
						PushButton 35, 250, 25, 10, "MEMI", MEMI_button
						PushButton 60, 250, 25, 10, "MONT", MONT_button
						PushButton 10, 260, 25, 10, "PARE", PARE_button
						PushButton 35, 260, 25, 10, "SANC", SANC_button
						PushButton 60, 260, 25, 10, "TIME", TIME_button
						PushButton 295, 245, 20, 10, "HC", ELIG_HC_button
						PushButton 295, 255, 20, 10, "MSA", ELIG_MSA_button
						PushButton 345, 245, 45, 10, "prev. panel", prev_panel_button
						PushButton 390, 245, 45, 10, "prev. memb", prev_memb_button
						PushButton 345, 255, 45, 10, "next panel", next_panel_button
						PushButton 390, 255, 45, 10, "next memb", next_memb_button
					Text 5, 15, 55, 10, "HRF datestamp:"
					Text 195, 15, 40, 10, "HRF status:"
					Text 5, 35, 45, 10, "Facility Info:"
					Text 230, 35, 50, 10, "Admit In Date:"
					Text 5, 55, 55, 10, "Earned income:"
					Text 5, 75, 60, 10, "Unearned income:"
					Text 5, 115, 30, 10, "Assets:"
					Text 5, 135, 40, 10, "Deductions:"
					Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
					Text 5, 175, 45, 10, "Other notes:"
					GroupBox 5, 190, 60, 40, "Income panels"
					GroupBox 70, 190, 110, 40, "Asset panels"
					GroupBox 5, 235, 85, 40, "other STAT panels:"
					Text 185, 195, 50, 10, "Verifs needed:"
					Text 185, 215, 50, 10, "Actions taken:"
					GroupBox 280, 230, 50, 40, "ELIG panels:"
					GroupBox 340, 230, 100, 40, "STAT-based navigation"
					Text 100, 280, 60, 10, "Worker signature:"
					If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
						GroupBox 90, 235, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
						Text 95, 245, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
						CheckBox 95, 255, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
					End If
					EndDialog
					err_msg = ""
					Dialog Dialog1
					cancel_confirmation
					Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
					MAXIS_dialog_navigation
					If ButtonPressed = income_notes_button Then
						'-------------------------------------------------------------------------------------------------DIALOG
						Dialog1 = "" 'Blanking out previous dialog detail
						BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
						CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
						CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
						CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
						CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
						CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
						CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
						CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
						CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
						CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
						CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
						CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
						ButtonGroup ButtonPressed
							PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
							CancelButton 295, 205, 50, 15
						Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
						GroupBox 5, 60, 340, 45, "JOBS Income"
						GroupBox 5, 110, 340, 45, "BUSI Income"
						GroupBox 5, 160, 340, 40, "UNEA Income"
						EndDialog
						Dialog Dialog1
						If ButtonPressed = add_to_notes_button Then
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
							If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
							If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
							If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
							If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
							If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
							If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
							If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
							If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
							If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
					IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
					IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
					IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
					IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
					IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
				case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
				IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
			LOOP UNTIL case_note_confirmation = vbYes
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false

		'Setting up some variables for the case note
		programs_list = "HC"
		If MSA_check = checked Then programs_list = programs_list & " & MSA"
		If admit_date <> "" then facility_info = facility_info & ". Admit Date: " & admit_date


		'Enters the case note-----------------------------------------------------------------------------------------------
		start_a_blank_CASE_NOTE
		Call write_variable_in_case_note("***" & MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
		call write_bullet_and_variable_in_case_note("Programs", programs_list)
		call write_bullet_and_variable_in_case_note("Facility", facility_info)
		call write_bullet_and_variable_in_case_note("Earned income", earned_income)
		call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
		call write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
		call write_bullet_and_variable_in_case_note("Assets", assets)
		call write_bullet_and_variable_in_case_note("Deductions", hc_deductions)
		call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
		call write_bullet_and_variable_in_case_note("Other notes", other_notes)
		If sent_3050_checkbox = 1 then call write_variable_in_CASE_NOTE("* Sent 3050 to Facility")
		If HRF_release_checkbox = 1 then call write_variable_in_CASE_NOTE("* Released HRF in MAXIS for next month.")
		IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
		call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
		call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
		If next_month_hrf_not_received_checkbox = checked Then
			call write_variable_in_CASE_NOTE("* HRF for next month (" & CM_plus_1_mo & "/" & CM_plus_1_yr & ") has not been received.")
			call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
		End If

		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)

		end_msg = "Success! Your HRF for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " on a LTC case has been case noted."
	ElseIf LTC_case = vbNo then							'Shows dialog if not LTC
		'The case note dialog, complete with panel navigation, reading the ELIG/MFIP screen, and navigation to case note, as well as logic for certain sections to be required.
		DO
			DO
				Do
					err_msg = ""
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					BeginDialog Dialog1, 0, 0, 451, 285, MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF dialog"
					EditBox 65, 30, 50, 15, HRF_datestamp
					DropListBox 170, 30, 75, 15, "complete"+chr(9)+"incomplete", HRF_status
					EditBox 65, 50, 380, 15, earned_income
					EditBox 70, 70, 375, 15, unearned_income
					EditBox 110, 90, 335, 15, notes_on_income
					EditBox 30, 110, 90, 15, YTD
					EditBox 170, 110, 275, 15, changes
					EditBox 30, 130, 415, 15, EMPS
					EditBox 100, 150, 345, 15, FIAT_reasons
					EditBox 50, 170, 395, 15, other_notes
					CheckBox 190, 190, 60, 10, "10% sanction?", ten_percent_sanction_check
					CheckBox 265, 190, 60, 10, "30% sanction?", thirty_percent_sanction_check
					CheckBox 330, 190, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
					EditBox 235, 205, 210, 15, verifs_needed
					EditBox 235, 225, 210, 15, actions_taken
					EditBox 340, 245, 105, 15, worker_signature
					ButtonGroup ButtonPressed
						OkButton 340, 265, 50, 15
						CancelButton 395, 265, 50, 15
						PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
						PushButton 260, 20, 20, 10, "FS", ELIG_FS_button
						PushButton 280, 20, 20, 10, "HC", ELIG_HC_button
						PushButton 300, 20, 20, 10, "MFIP", ELIG_MFIP_button
						PushButton 260, 30, 20, 10, "GA", ELIG_GA_button
						PushButton 335, 20, 45, 10, "prev. panel", prev_panel_button
						PushButton 395, 20, 45, 10, "prev. memb", prev_memb_button
						PushButton 335, 30, 45, 10, "next panel", next_panel_button
						PushButton 395, 30, 45, 10, "next memb", next_memb_button
						PushButton 5, 135, 25, 10, "EMPS", EMPS_button
						PushButton 10, 205, 25, 10, "BUSI", BUSI_button
						PushButton 35, 205, 25, 10, "JOBS", JOBS_button
						PushButton 10, 215, 25, 10, "RBIC", RBIC_button
						PushButton 35, 215, 25, 10, "UNEA", UNEA_button
						PushButton 75, 205, 25, 10, "ACCT", ACCT_button
						PushButton 100, 205, 25, 10, "CARS", CARS_button
						PushButton 125, 205, 25, 10, "CASH", CASH_button
						PushButton 150, 205, 25, 10, "OTHR", OTHR_button
						PushButton 75, 215, 25, 10, "REST", REST_button
						PushButton 100, 215, 25, 10, "SECU", SECU_button
						PushButton 125, 215, 25, 10, "TRAN", TRAN_button
						PushButton 10, 250, 25, 10, "MEMB", MEMB_button
						PushButton 35, 250, 25, 10, "MEMI", MEMI_button
						PushButton 60, 250, 25, 10, "MONT", MONT_button
						PushButton 10, 260, 25, 10, "PARE", PARE_button
						PushButton 35, 260, 25, 10, "SANC", SANC_button
						PushButton 60, 260, 25, 10, "TIME", TIME_button
					Text 5, 115, 20, 10, "YTD:"
					Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
					Text 5, 175, 45, 10, "Other notes:"
					GroupBox 5, 190, 60, 40, "Income panels"
					GroupBox 70, 190, 110, 40, "Asset panels"
					Text 280, 250, 60, 10, "Worker signature:"
					Text 185, 230, 50, 10, "Actions taken:"
					GroupBox 5, 235, 85, 40, "other STAT panels:"
					Text 185, 210, 50, 10, "Verifs needed:"
					Text 125, 35, 40, 10, "HRF status:"
					Text 130, 115, 35, 10, "Changes?:"
					GroupBox 330, 5, 115, 40, "STAT-based navigation"
					Text 5, 35, 55, 10, "HRF datestamp:"
					Text 5, 55, 55, 10, "Earned income:"
					Text 5, 75, 60, 10, "Unearned income:"
					GroupBox 255, 5, 70, 40, "ELIG panels:"
					If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
						GroupBox 90, 245, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
						Text 95, 255, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
						CheckBox 95, 265, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
					End If

					EndDialog
					Dialog Dialog1
					cancel_confirmation
					Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
					MAXIS_dialog_navigation
					If ButtonPressed = income_notes_button Then
						'-------------------------------------------------------------------------------------------------DIALOG
						Dialog1 = "" 'Blanking out previous dialog detail
						BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
						CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
						CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
						CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
						CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
						CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
						CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
						CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
						CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
						CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
						CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
						CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
						ButtonGroup ButtonPressed
							PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
							CancelButton 295, 205, 50, 15
						Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
						GroupBox 5, 60, 340, 45, "JOBS Income"
						GroupBox 5, 110, 340, 45, "BUSI Income"
						GroupBox 5, 160, 340, 40, "UNEA Income"
						EndDialog
						Dialog dialog1
						If ButtonPressed = add_to_notes_button Then
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
							If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
							If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
							If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
							If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
							If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
							If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
							If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
							If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
							If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
					IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
					IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
					IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
					IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
					IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
				case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
				IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
			LOOP UNTIL case_note_confirmation = vbYes
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false

		'Creating program list---------------------------------------------------------------------------------------------
		If MFIP_check = 1 Then programs_list = "MFIP "
		If SNAP_check = 1 Then programs_list = programs_list & "SNAP "
		If HC_check = 1 Then programs_list = programs_list & "HC "
		If GA_check = 1 Then programs_list = programs_list & "GA "
		If MSA_check = 1 Then programs_list = programs_list & "MSA "

		'Enters the case note-----------------------------------------------------------------------------------------------
		start_a_blank_CASE_NOTE
		Call write_variable_in_case_note("***" & HRF_month & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
		call write_bullet_and_variable_in_case_note("Programs", programs_list)
		call write_bullet_and_variable_in_case_note("Earned income", earned_income)
		call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
		CALL write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
		call write_bullet_and_variable_in_case_note("YTD", YTD)
		call write_bullet_and_variable_in_case_note("Changes", changes)
		call write_bullet_and_variable_in_case_note("EMPS", EMPS)
		call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
		call write_bullet_and_variable_in_case_note("Other notes", other_notes)
		If ten_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 10% sanction.")
		If thirty_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 30% sanction.")
		IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
		call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
		call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
		If next_month_hrf_not_received_checkbox = checked Then
			call write_variable_in_CASE_NOTE("* HRF for next month (" & next_HRF_month & ") has not been received.")
			call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
		End If
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)



		end_msg = "Success! Your HRF for " & HRF_month & " has been case noted."

	End If

	script_end_procedure_with_error_report(end_msg & vbcr & "Please make sure to accept the Work items in ECF associated with this HRF. Thank you!")

Else
	'Benefit month equal to 03 25 or after

	'NAV to STAT
	call navigate_to_MAXIS_screen("STAT", "MEMB")

	'Creating a custom dialog for determining who the HH members are
	call HH_member_custom_dialog(HH_member_array)

	'To do - handling for GA - GRH PIC and UNEA - no clear instructions
	'Autofilling info for case note

	'If it is a MFIP or SNAP case then use new JOBS panel reading
	If instr(programs_selected_list, "MFIP") or instr(programs_selected_list, "SNAP") Then
		call HRF_autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
	Else
		call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
	End If
	call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
	call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
	'If it is a MFIP or SNAP case then use new UNEA panel reading
	If instr(programs_selected_list, "MFIP") or instr(programs_selected_list, "SNAP") Then
		call HRF_autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
	Else
		call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
	End If
	If instr(programs_selected_list, "SNAP") Then
		call HRF_autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
	Else
		call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
	End If
	call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)

	'Cleaning up info for case note
	HRF_computer_friendly_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
	retro_month_name = monthname(datepart("m", (dateadd("m", -2, HRF_computer_friendly_month))))
	pro_month_name = monthname(datepart("m", (HRF_computer_friendly_month)))
	HRF_month = retro_month_name & "/" & pro_month_name
	next_month_hrf_not_received_checkbox = unchecked
	next_retro_month_name = monthname(datepart("m", (dateadd("m", -1, HRF_computer_friendly_month))))
	next_month_name = monthname(datepart("m", DateAdd("m", 1, HRF_computer_friendly_month)))
	next_HRF_month = next_retro_month_name & "/" & next_month_name

	'If a HRF is being run for a HC case, script will ask if this is a LTC case
	If HC_check = checked Then
		'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
		LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have different fields in their dialog.", vbYesNoCancel)
		If LTC_case = vbCancel then stopscript
	Else
		LTC_case = vbNo
	End If

	'If workers answers yes to this is a LTC case - script runs this specific functionality
	If LTC_case = vbYes then
		'LTC cases should not have these programs active
		If MFIP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* MFIP will be removed."
		If SNAP_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* SNAP will be removed."
		If GA_check = checked Then uncheck_msg = uncheck_msg & vbNewLine & "* GA will be removed."

		'Alerting worker that these programs will be unchecked.
		If uncheck_msg <> "" Then MsgBox "You have checked programs that should not be active with LTC. These programs will not be added to the note." & vbNewLine & uncheck_msg

		MFIP_check = unchecked
		SNAP_check = unchecked
		GA_check = unchecked

		'Getting some additional information for the dialog to be autofilled
		call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
		call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
		call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)

		'Going to find the current facility to autofil the dialog
		Call navigate_to_MAXIS_screen ("STAT", "FACI")

		'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
		Do
			EMReadScreen FACI_current_panel, 1, 2, 73
			EMReadScreen FACI_total_check, 1, 2, 78
			EMReadScreen in_year_check_01, 4, 14, 53
			EMReadScreen in_year_check_02, 4, 15, 53
			EMReadScreen in_year_check_03, 4, 16, 53
			EMReadScreen in_year_check_04, 4, 17, 53
			EMReadScreen in_year_check_05, 4, 18, 53
			EMReadScreen out_year_check_01, 4, 14, 77
			EMReadScreen out_year_check_02, 4, 15, 77
			EMReadScreen out_year_check_03, 4, 16, 77
			EMReadScreen out_year_check_04, 4, 17, 77
			EMReadScreen out_year_check_05, 4, 18, 77
			If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or (in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
				currently_in_FACI = True
				exit do
			Elseif FACI_current_panel = FACI_total_check then
				currently_in_FACI = False
				exit do
			Else
				transmit
			End if
		Loop until FACI_current_panel = FACI_total_check

		If in_year_check_01 <> "____" and out_year_check_01 = "____" Then EMReadScreen date_in, 10, 14, 47
		If in_year_check_02 <> "____" and out_year_check_02 = "____" Then EMReadScreen date_in, 10, 15, 47
		If in_year_check_03 <> "____" and out_year_check_03 = "____" Then EMReadScreen date_in, 10, 16, 47
		If in_year_check_04 <> "____" and out_year_check_04 = "____" Then EMReadScreen date_in, 10, 17, 47
		If in_year_check_05 <> "____" and out_year_check_05 = "____" Then EMReadScreen date_in, 10, 18, 47

		admit_date = replace(date_in, " ", "/")

		'Gets Facility name and admit in date and enters it into the dialog
		If currently_in_FACI = True then
			EMReadScreen FACI_name, 30, 6, 43
			facility_info = trim(replace(FACI_name, "_", ""))
		End if

		'confirms that case is in the footer month/year selected by the user
		Call MAXIS_footer_month_confirmation
		Call MAXIS_background_check

		'Goes to STAT WKEX to get deductions and possible FIAT reasons to autofil the dialog
		Call navigate_to_MAXIS_screen("STAT", "WKEX")
		EMReadScreen WKEX_check, 1, 2, 73
		If WKEX_check = "0" then
			script_end_procedure("You do not have a WKEX panel. Please create a WKEX panel for your HC case, and re-run the script.")
		Elseif WKEX_check <> "0" then
			'Reads work expenses for MEMB 01, verification codes and impairment related code
			EMReadScreen program_check,          2, 5, 33
			EMReadScreen federal_tax,            8, 7, 57
			EMReadScreen federal_tax_verif_code, 1, 7, 69
			EMReadScreen state_tax,              8, 8, 57
			EMReadScreen state_tax_verif_code, 1, 8, 69
			EMReadScreen FICA_witheld,           8, 9, 57
			EMReadScreen FICA_witheld_verif_code, 1, 9, 69
			EMReadScreen transportation_expense, 8, 10, 57
			EMReadScreen transportation_expense_verif_code, 1, 10, 69
			EMReadScreen transportation_impair, 1, 10, 75
			EMReadScreen meals_expense, 8, 11, 57
			EMReadScreen meals_impair, 1, 11, 75
			EMReadScreen meals_expense_verif_code, 1, 11, 69
			EMReadScreen uniform_expense, 8, 12, 57
			EMReadScreen uniform_expense_verif_code, 1, 12, 69
			EMReadScreen uniform_impair, 1, 12, 75
			EMReadScreen tools_expense, 8, 13, 57
			EMReadScreen tools_expense_verif_code, 1, 13, 69
			EMReadScreen tools_impair, 1, 13, 75
			EMReadScreen dues_expense, 8, 14, 57
			EMReadScreen dues_expense_verif_code, 1, 14, 69
			EMReadScreen dues_impair, 1, 14, 75
			EMReadScreen other_expense, 8, 15,	57
			EMReadScreen other_expense_verif_code, 1, 15, 69
			EMReadScreen other_impair, 1, 15, 75
		End IF

		'cleaning up the WKEX variables
		federal_tax = replace(federal_tax, "_", "")
		federal_tax = trim(federal_tax)
		state_tax = replace(state_tax, "_", "")
		state_tax = trim(state_tax)
		FICA_witheld = replace(FICA_witheld, "_", "")
		FICA_witheld = trim(FICA_witheld)
		transportation_expense = replace(transportation_expense, "_", "")
		transportation_expense = trim(transportation_expense)
		meals_expense = replace(meals_expense, "_", "")
		meals_expense = trim(meals_expense)
		uniform_expense = replace(uniform_expense, "_", "")
		uniform_expense = trim(uniform_expense)
		tools_expense = replace(tools_expense, "_", "")
		tools_expense = trim(tools_expense)
		dues_expense = replace(dues_expense, "_", "")
		dues_expense = trim(dues_expense)
		other_expense = replace(other_expense, "_", "")
		other_expense = trim(other_expense)

		'Gives unverified expenses and blank expenses the value of $0 and adds non-zero amounts to the dialog for autofil
		If federal_tax = "" OR federal_tax_verif_code = "N" then
			federal_tax = "0"
		Else
			hc_deductions = hc_deductions & "; Federal Tax - $" & federal_tax
		End if
		If state_tax = "" OR state_tax_verif_code = "N" then
			state_tax = "0"
		Else
			hc_deductions = hc_deductions & "; State Tax - $" & state_tax
		End if
		If FICA_witheld = "" OR FICA_witheld_verif_code = "N" then
			FICA_witheld = "0"
		Else
			hc_deductions = hc_deductions & "; FICA - $" & FICA_witheld
		End if
		If transportation_expense = "" OR transportation_expense_verif_code = "N" OR transportation_impair =  "_" OR transportation_impair = "N" then
			transportation_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Transportation Expense - $" & transportation_expense
		End if
		If meals_expense = "" OR meals_expense_verif_code = "N" OR meals_impair = "_" OR meals_impair = "N" then
			meals_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Meals Expense - $" & meals_expense
		End if
		If uniform_expense = "" OR uniform_expense_verif_code = "N" OR uniform_impair = "_" OR uniform_impair = "N" then
			uniform_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Uniform Expense - $" & uniform_expense
		End if
		If tools_expense = "" OR tools_expense_verif_code = "N" OR tools_impair = "_" OR tools_impair = "N" then
			tools_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Tools Expense - $" & tools_expense
		End if
		If dues_expense = "" OR dues_expense_verif_code = "N" OR dues_impair = "_" OR dues_impair = "N" then
			dues_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Dues Expense - $" & dues_expense
		End if
		If other_expense = "" OR other_expense_verif_code = "N" OR other_impair = "_" OR other_impair = "N" then
			other_expense = "0"
		Else
			hc_deductions = hc_deductions & "; Other Expense - $" & other_expense
		End if

		'Checks PDED other expenses, will need to add PDED and WKEX other expenses together
		Call navigate_to_MAXIS_screen("STAT", "PDED")
		EMReadScreen other_earned_income_PDED, 8, 11, 62

		'cleaning up PDED variables
		other_earned_income_PDED = replace(other_earned_income_PDED, "_", "")
		other_earned_income_PDED = trim(other_earned_income_PDED)

		'Gives blank expenses the value of $0
		If other_earned_income_PDED = "" then
			other_earned_income_PDED = "0"
		Else
			hc_deductions = hc_deductions & "; Other Earned Income Deductions - $" & other_earned_income_PDED
		End If

		'Determining if earned income is less than $80
		Call navigate_to_MAXIS_screen ("STAT", "JOBS")
		EMReadScreen JOBS_panel_income, 7, 17, 68
		JOBS_panel_income = trim(JOBS_panel_income)
		If IsNumeric(JOBS_panel_income) = TRUE Then
			If abs(JOBS_panel_income) < 80 then
				special_pers_allow = JOBS_panel_income	'if less then $80 deduction is earned income amount
			ELSE
				special_pers_allow = "80.00"		'otherwise deduction is $80
			END IF
		Else
			JOBS_panel_income = ""
		End If

		If JOBS_panel_income <> "" Then hc_deductions = hc_deductions & "; Special Allowance - $" & special_pers_allow

		'All of the deductions found to this point need to be FIATed. Added these to the FIAT varirable.
		FIAT_reasons = hc_deductions

		'Going to see if there is a deduction on MEDI. (This does not have to be FIATED)
		Call navigate_to_MAXIS_screen ("STAT", "MEDI")
		EMReadScreen medi_panel_exists, 1, 2, 78
		If medi_panel_exists = "1" Then
			EMReadScreen part_b_premium, 9, 7, 72
			part_b_premium = trim(part_b_premium)
			If part_b_premium <> "________" Then hc_deductions = hc_deductions & "; Medicare Premium - $" & part_b_premium
		End If

		'Formatting the variables for the dialog
		hc_deductions = right(hc_deductions, len(hc_deductions) - 2)
		FIAT_reasons = right(FIAT_reasons, len(FIAT_reasons) - 2)

	
		'The case note dialog, complete with panel navigation, reading the ELIG/MSA or ELIG/HC screen, and navigation to case note, as well as logic for certain sections to be required.
		DO
			DO
				Do
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					BeginDialog Dialog1, 0, 0, 451, 295, "HRF for LTC Cases"
					EditBox 65, 10, 85, 15, HRF_datestamp
					DropListBox 240, 10, 80, 15, "complete"+chr(9)+"incomplete", HRF_status
					EditBox 50, 30, 165, 15, facility_info
					EditBox 280, 30, 55, 15, admit_date
					CheckBox 350, 5, 80, 10, "Sent 3050 to Facility", sent_3050_checkbox
					CheckBox 350, 20, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
					CheckBox 350, 35, 80, 10, "Next HRF Released", HRF_release_checkbox
					EditBox 65, 50, 380, 15, earned_income
					EditBox 70, 70, 375, 15, unearned_income
					EditBox 110, 90, 335, 15, notes_on_income
					EditBox 40, 110, 405, 15, assets
					EditBox 50, 130, 395, 15, hc_deductions
					EditBox 100, 150, 345, 15, FIAT_reasons
					EditBox 50, 170, 395, 15, other_notes
					EditBox 235, 190, 210, 15, verifs_needed
					EditBox 235, 210, 210, 15, actions_taken
					EditBox 165, 275, 105, 15, worker_signature
					ButtonGroup ButtonPressed
						OkButton 340, 275, 50, 15
						CancelButton 390, 275, 50, 15
						PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
						PushButton 10, 205, 25, 10, "BUSI", BUSI_button
						PushButton 35, 205, 25, 10, "JOBS", JOBS_button
						PushButton 10, 215, 25, 10, "RBIC", RBIC_button
						PushButton 35, 215, 25, 10, "UNEA", UNEA_button
						PushButton 75, 205, 25, 10, "ACCT", ACCT_button
						PushButton 100, 205, 25, 10, "CARS", CARS_button
						PushButton 125, 205, 25, 10, "CASH", CASH_button
						PushButton 150, 205, 25, 10, "OTHR", OTHR_button
						PushButton 75, 215, 25, 10, "REST", REST_button
						PushButton 100, 215, 25, 10, "SECU", SECU_button
						PushButton 125, 215, 25, 10, "TRAN", TRAN_button
						PushButton 10, 250, 25, 10, "MEMB", MEMB_button
						PushButton 35, 250, 25, 10, "MEMI", MEMI_button
						PushButton 60, 250, 25, 10, "MONT", MONT_button
						PushButton 10, 260, 25, 10, "PARE", PARE_button
						PushButton 35, 260, 25, 10, "SANC", SANC_button
						PushButton 60, 260, 25, 10, "TIME", TIME_button
						PushButton 295, 245, 20, 10, "HC", ELIG_HC_button
						PushButton 295, 255, 20, 10, "MSA", ELIG_MSA_button
						PushButton 345, 245, 45, 10, "prev. panel", prev_panel_button
						PushButton 390, 245, 45, 10, "prev. memb", prev_memb_button
						PushButton 345, 255, 45, 10, "next panel", next_panel_button
						PushButton 390, 255, 45, 10, "next memb", next_memb_button
					Text 5, 15, 55, 10, "HRF datestamp:"
					Text 195, 15, 40, 10, "HRF status:"
					Text 5, 35, 45, 10, "Facility Info:"
					Text 230, 35, 50, 10, "Admit In Date:"
					Text 5, 55, 55, 10, "Earned income:"
					Text 5, 75, 60, 10, "Unearned income:"
					Text 5, 115, 30, 10, "Assets:"
					Text 5, 135, 40, 10, "Deductions:"
					Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
					Text 5, 175, 45, 10, "Other notes:"
					GroupBox 5, 190, 60, 40, "Income panels"
					GroupBox 70, 190, 110, 40, "Asset panels"
					GroupBox 5, 235, 85, 40, "other STAT panels:"
					Text 185, 195, 50, 10, "Verifs needed:"
					Text 185, 215, 50, 10, "Actions taken:"
					GroupBox 280, 230, 50, 40, "ELIG panels:"
					GroupBox 340, 230, 100, 40, "STAT-based navigation"
					Text 100, 280, 60, 10, "Worker signature:"
					If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
						GroupBox 90, 235, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
						Text 95, 245, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
						CheckBox 95, 255, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
					End If
					EndDialog
					err_msg = ""
					Dialog Dialog1
					cancel_confirmation
					Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
					MAXIS_dialog_navigation
					If ButtonPressed = income_notes_button Then
						'-------------------------------------------------------------------------------------------------DIALOG
						Dialog1 = "" 'Blanking out previous dialog detail
						BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
						CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
						CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
						CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
						CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
						CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
						CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
						CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
						CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
						CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
						CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
						CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
						ButtonGroup ButtonPressed
							PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
							CancelButton 295, 205, 50, 15
						Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
						GroupBox 5, 60, 340, 45, "JOBS Income"
						GroupBox 5, 110, 340, 45, "BUSI Income"
						GroupBox 5, 160, 340, 40, "UNEA Income"
						EndDialog
						Dialog Dialog1
						If ButtonPressed = add_to_notes_button Then
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
							If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
							If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
							If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
							If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
							If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
							If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
							If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
							If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
							If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
					IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
					IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
					IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
					IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
					IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
				case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
				IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
			LOOP UNTIL case_note_confirmation = vbYes
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false

		'Setting up some variables for the case note
		programs_list = "HC"
		If MSA_check = checked Then programs_list = programs_list & " & MSA"
		If admit_date <> "" then facility_info = facility_info & ". Admit Date: " & admit_date


		'Enters the case note-----------------------------------------------------------------------------------------------
		start_a_blank_CASE_NOTE
		Call write_variable_in_case_note("***" & MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
		call write_bullet_and_variable_in_case_note("Programs", programs_list)
		call write_bullet_and_variable_in_case_note("Facility", facility_info)
		call write_bullet_and_variable_in_case_note("Earned income", earned_income)
		call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
		call write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
		call write_bullet_and_variable_in_case_note("Assets", assets)
		call write_bullet_and_variable_in_case_note("Deductions", hc_deductions)
		call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
		call write_bullet_and_variable_in_case_note("Other notes", other_notes)
		If sent_3050_checkbox = 1 then call write_variable_in_CASE_NOTE("* Sent 3050 to Facility")
		If HRF_release_checkbox = 1 then call write_variable_in_CASE_NOTE("* Released HRF in MAXIS for next month.")
		IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
		call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
		call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
		If next_month_hrf_not_received_checkbox = checked Then
			call write_variable_in_CASE_NOTE("* HRF for next month (" & CM_plus_1_mo & "/" & CM_plus_1_yr & ") has not been received.")
			call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
		End If

		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)

		end_msg = "Success! Your HRF for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " on a LTC case has been case noted."
	ElseIf LTC_case = vbNo then							'Shows dialog if not LTC
		'The case note dialog, complete with panel navigation, reading the ELIG/MFIP screen, and navigation to case note, as well as logic for certain sections to be required.

		'Setting initial other notes text
		other_notes = "This is the prospective amount for the 6-month budget."
		DO
			DO
				Do
					err_msg = ""
					'-------------------------------------------------------------------------------------------------DIALOG
					Dialog1 = "" 'Blanking out previous dialog detail
					BeginDialog Dialog1, 0, 0, 451, 285, MAXIS_footer_month & "/" & MAXIS_footer_year & " HRF dialog"
					EditBox 65, 30, 50, 15, HRF_datestamp
					DropListBox 170, 30, 75, 15, "complete"+chr(9)+"incomplete", HRF_status
					EditBox 65, 50, 380, 15, earned_income
					EditBox 70, 70, 375, 15, unearned_income
					EditBox 110, 90, 335, 15, notes_on_income
					EditBox 30, 110, 90, 15, YTD
					EditBox 170, 110, 275, 15, changes
					EditBox 30, 130, 415, 15, EMPS
					EditBox 100, 150, 345, 15, FIAT_reasons
					EditBox 50, 170, 395, 15, other_notes
					CheckBox 190, 190, 60, 10, "10% sanction?", ten_percent_sanction_check
					CheckBox 265, 190, 60, 10, "30% sanction?", thirty_percent_sanction_check
					CheckBox 330, 190, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
					EditBox 235, 205, 210, 15, verifs_needed
					EditBox 235, 225, 210, 15, actions_taken
					EditBox 340, 245, 105, 15, worker_signature
					ButtonGroup ButtonPressed
						OkButton 340, 265, 50, 15
						CancelButton 395, 265, 50, 15
						PushButton 5, 95, 100, 10, "Notes on Income and Budget", income_notes_button
						PushButton 260, 20, 20, 10, "FS", ELIG_FS_button
						PushButton 280, 20, 20, 10, "HC", ELIG_HC_button
						PushButton 300, 20, 20, 10, "MFIP", ELIG_MFIP_button
						PushButton 260, 30, 20, 10, "GA", ELIG_GA_button
						PushButton 335, 20, 45, 10, "prev. panel", prev_panel_button
						PushButton 395, 20, 45, 10, "prev. memb", prev_memb_button
						PushButton 335, 30, 45, 10, "next panel", next_panel_button
						PushButton 395, 30, 45, 10, "next memb", next_memb_button
						PushButton 5, 135, 25, 10, "EMPS", EMPS_button
						PushButton 10, 205, 25, 10, "BUSI", BUSI_button
						PushButton 35, 205, 25, 10, "JOBS", JOBS_button
						PushButton 10, 215, 25, 10, "RBIC", RBIC_button
						PushButton 35, 215, 25, 10, "UNEA", UNEA_button
						PushButton 75, 205, 25, 10, "ACCT", ACCT_button
						PushButton 100, 205, 25, 10, "CARS", CARS_button
						PushButton 125, 205, 25, 10, "CASH", CASH_button
						PushButton 150, 205, 25, 10, "OTHR", OTHR_button
						PushButton 75, 215, 25, 10, "REST", REST_button
						PushButton 100, 215, 25, 10, "SECU", SECU_button
						PushButton 125, 215, 25, 10, "TRAN", TRAN_button
						PushButton 10, 250, 25, 10, "MEMB", MEMB_button
						PushButton 35, 250, 25, 10, "MEMI", MEMI_button
						PushButton 60, 250, 25, 10, "MONT", MONT_button
						PushButton 10, 260, 25, 10, "PARE", PARE_button
						PushButton 35, 260, 25, 10, "SANC", SANC_button
						PushButton 60, 260, 25, 10, "TIME", TIME_button
					Text 5, 115, 20, 10, "YTD:"
					Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
					Text 5, 175, 45, 10, "Other notes:"
					GroupBox 5, 190, 60, 40, "Income panels"
					GroupBox 70, 190, 110, 40, "Asset panels"
					Text 280, 250, 60, 10, "Worker signature:"
					Text 185, 230, 50, 10, "Actions taken:"
					GroupBox 5, 235, 85, 40, "other STAT panels:"
					Text 185, 210, 50, 10, "Verifs needed:"
					Text 125, 35, 40, 10, "HRF status:"
					Text 130, 115, 35, 10, "Changes?:"
					GroupBox 330, 5, 115, 40, "STAT-based navigation"
					Text 5, 35, 55, 10, "HRF datestamp:"
					Text 5, 55, 55, 10, "Earned income:"
					Text 5, 75, 60, 10, "Unearned income:"
					GroupBox 255, 5, 70, 40, "ELIG panels:"
					If CM_mo = MAXIS_footer_month and CM_yr = MAXIS_footer_year Then
						GroupBox 90, 245, 190, 35, "HRF BEING PROCESSED IN THE BENEFIT MONTH."
						Text 95, 255, 180, 10, "This means there may be a HRF due for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " as well."
						CheckBox 95, 265, 180, 10, "Check here if HRF for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " is due and not received.", next_month_hrf_not_received_checkbox
					End If

					EndDialog
					Dialog Dialog1
					cancel_confirmation
					Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
					MAXIS_dialog_navigation
					If ButtonPressed = income_notes_button Then
						'-------------------------------------------------------------------------------------------------DIALOG
						Dialog1 = "" 'Blanking out previous dialog detail
						BeginDialog Dialog1, 0, 0, 351, 225, "Explanation of Income"
						CheckBox 10, 25, 300, 10, "NO INCOME - All income previously ended, but case has not yet fallen off of the HRF run. ", no_income_checkbox
						CheckBox 10, 35, 225, 10, "SEE PREVIOUS NOTE - Income detail is listed on previous note(s)", see_other_note_checkbox
						CheckBox 10, 45, 280, 10, "NOT VERIFIED - Income has not been fully verified, detail will be entered in future.", not_verified_checkbox
						CheckBox 10, 70, 285, 10, "VERIFS RECEIVED - All pay verification provided for the report month and budgeted.", jobs_all_verified_checkbox
						CheckBox 10, 80, 260, 10, "PARTIAL MONTH - Job ended in the report month, all pay has been budgeted.", jobs_partial_month_checkbox
						CheckBox 10, 90, 305, 10, "YEAR TO DATE USED - Not all pay dates verified, check amount was calculated using YTD.", jobs_ytd_used_checkbox
						CheckBox 10, 120, 320, 10, "SELF EMP REPORT FORM - DHS 3336 submitted as proof of monthly self employment income.", busi_rept_form_checkbox
						CheckBox 10, 130, 310, 10, "EXPENSES - 50% - Budgeted self emp. income - allowing 50% of gross income as expenses.", busi_fifty_percent_checkbox
						CheckBox 10, 140, 235, 10, "TAX METHOD - Self employment income budgeted using tax method.", busi_tax_method_checkbox
						CheckBox 10, 175, 270, 10, "VERIFS RECEIVED - All verification of unearned income in report month received.", unea_all_verified_checkbox
						CheckBox 10, 185, 320, 10, "UNCHANGING - Unearned income does not vary and no change reported for this report month.", unea_unvarying_checkbox
						ButtonGroup ButtonPressed
							PushButton 240, 205, 50, 15, "Insert", add_to_notes_button
							CancelButton 295, 205, 50, 15
						Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
						GroupBox 5, 60, 340, 45, "JOBS Income"
						GroupBox 5, 110, 340, 45, "BUSI Income"
						GroupBox 5, 160, 340, 40, "UNEA Income"
						EndDialog
						Dialog dialog1
						If ButtonPressed = add_to_notes_button Then
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; All income ended prior to " & retro_month_name & " - no income budgeted."
							If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
							If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all pay dates in " & retro_month_name & "."
							If jobs_partial_month_checkbox = checked Then notes_on_income = notes_on_income & "; Job ended in " & retro_month_name & " and pay was not received on every pay date. All income received has been budgeted."
							If jobs_ytd_used_checkbox = checked Then notes_on_income = notes_on_income & "; Not all pay date amounts were verified, able to use year to date amounts to calculate missing pay date amounts."
							If busi_rept_form_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided Self Employment Report Form (DHS 3336) to verify self employment income in " & retro_month_name
							If busi_fifty_percent_checkbox = checked Then notes_on_income = notes_on_income & "; Gross self employment income from " & retro_month_name & " verified - allowed 50% of this gross amount as an expense."
							If busi_tax_method_checkbox = checked Then notes_on_income = notes_on_income & "; Self Employment income budgeted using TAX records."
							If unea_all_verified_checkbox = checked Then notes_on_income = notes_on_income & "; Client provided verification of all unearned income in " & retro_month_name & "."
							If unea_unvarying_checkbox = checked Then notes_on_income = notes_on_income & "; Unearned income on this case is unvarying."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
					IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
					IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
					IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
					IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
					IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
				case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
				IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
			LOOP UNTIL case_note_confirmation = vbYes
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false

		'Creating program list---------------------------------------------------------------------------------------------
		If MFIP_check = 1 Then programs_list = "MFIP "
		If SNAP_check = 1 Then programs_list = programs_list & "SNAP "
		If HC_check = 1 Then programs_list = programs_list & "HC "
		If GA_check = 1 Then programs_list = programs_list & "GA "
		If MSA_check = 1 Then programs_list = programs_list & "MSA "

		'Enters the case note-----------------------------------------------------------------------------------------------
		start_a_blank_CASE_NOTE
		Call write_variable_in_case_note("***6 month budget conversion, " & HRF_month & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
		call write_bullet_and_variable_in_case_note("Programs", programs_list)
		call write_bullet_and_variable_in_case_note("Earned income", earned_income)
		call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
		CALL write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
		call write_bullet_and_variable_in_case_note("YTD", YTD)
		call write_bullet_and_variable_in_case_note("Changes", changes)
		call write_bullet_and_variable_in_case_note("EMPS", EMPS)
		call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
		call write_bullet_and_variable_in_case_note("Other notes", other_notes)
		If ten_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 10% sanction.")
		If thirty_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 30% sanction.")
		IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
		call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
		call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
		If next_month_hrf_not_received_checkbox = checked Then
			call write_variable_in_CASE_NOTE("* HRF for next month (" & next_HRF_month & ") has not been received.")
			call write_variable_in_CASE_NOTE("  This may cause closure for " & CM_plus_1_mo & "/" & CM_plus_1_yr & " if not received.")
		End If
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)

		end_msg = "Success! Your HRF for " & HRF_month & " has been case noted."

	End If

	script_end_procedure_with_error_report(end_msg & vbcr & "Please make sure to accept the Work items in ECF associated with this HRF. Thank you!")
End If


'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------01/31/2025
'--Tab orders reviewed & confirmed----------------------------------------------01/31/2025
'--Mandatory fields all present & Reviewed--------------------------------------01/31/2025
'--All variables in dialog match mandatory fields-------------------------------01/31/2025
'Review dialog names for content and content fit in dialog----------------------01/31/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------01/31/2025
'--CASE:NOTE Header doesn't look funky------------------------------------------01/31/2025
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------01/31/2025
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----01/31/2025
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------01/31/2025
'--MAXIS_background_check reviewed (if applicable)------------------------------01/31/2025
'--PRIV Case handling reviewed -------------------------------------------------01/31/2025
'--Out-of-County handling reviewed----------------------------------------------NA
'--script_end_procedures (w/ or w/o error messaging)----------------------------01/31/2025
'--BULK - review output of statistics and run time/count (if applicable)--------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------01/31/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------01/31/2025
'--Incrementors reviewed (if necessary)-----------------------------------------01/31/2025
'--Denomination reviewed -------------------------------------------------------01/31/2025
'--Script name reviewed---------------------------------------------------------01/31/2025
'--BULK - remove 1 incrementor at end of script reviewed------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------01/31/2025
'--comment Code-----------------------------------------------------------------01/31/2025
'--Update Changelog for release/update------------------------------------------01/31/2025
'--Remove testing message boxes-------------------------------------------------01/31/2025
'--Remove testing code/unnecessary code-----------------------------------------01/31/2025
'--Review/update SharePoint instructions----------------------------------------01/31/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------01/31/2025
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------01/31/2025
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------01/31/2025
'--Complete misc. documentation (if applicable)---------------------------------01/31/2025
'--Update project team/issue contact (if applicable)----------------------------01/31/2025