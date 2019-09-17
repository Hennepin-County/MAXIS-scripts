
'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 720                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation= ucase(windows_user_ID)
name_for_validation = ""

If user_ID_for_validation = "CALO001" Then name_for_validation = "Casey"
If user_ID_for_validation = "ILFE001" Then name_for_validation = "Ilse"
If user_ID_for_validation = "WFS395" Then name_for_validation = "MiKayla"

If user_ID_for_validation = "FLWI002" Then name_for_validation = "Florence"
If user_ID_for_validation = "AAGA001" Then name_for_validation = "Aaron"
If user_ID_for_validation = "WFC041" Then name_for_validation = "Kerry"
If user_ID_for_validation = "TAPA002" Then name_for_validation = "Tanya"
If user_ID_for_validation = "AIRO001" Then name_for_validation = "Aimee"
If user_ID_for_validation = "KESE001" Then name_for_validation = "Keith"
If user_ID_for_validation = "WFQ898" Then name_for_validation = "Hannah"
If user_ID_for_validation = "WFU161" Then name_for_validation = "Brooke"
If user_ID_for_validation = "WFP106" Then name_for_validation = "Deborah"
If user_ID_for_validation = "WFK093" Then name_for_validation = "Jessica"
If user_ID_for_validation = "WFM207" Then name_for_validation = "Mandora"
If user_ID_for_validation = "WFI021" Then name_for_validation = "Brenda"
If user_ID_for_validation = "JAAR001" Then name_for_validation = "Jacob"
' If user_ID_for_validation = "REHU001" Then name_for_validation = "Remy"
If user_ID_for_validation = "WFI438" Then name_for_validation = "Leah"
If user_ID_for_validation = "WFJ018" Then name_for_validation = "Tracy"
If user_ID_for_validation = "AMKE001" Then name_for_validation = "Amy"
If user_ID_for_validation = "WFQ610" Then name_for_validation = "Kary"
If user_ID_for_validation = "WFP430" Then name_for_validation = "Amorette"
If user_ID_for_validation = "LACH001" Then name_for_validation = "Lara"

If name_for_validation <> "" Then
    ' MsgBox "Hello " & name_for_validation &  ", you have been selected to test the script NOTES - CAF."  & vbNewLine & vbNewLine & "A testing version of the script will now run.  Thank you for taking your time to review our new scripts and functionality as we strive for Continuous Improvement." & vbNewLine & vbNewLine  & "                                                                                    - BlueZone Script Team"
    ' testing_run = TRUE
    If run_locally = true then
        testing_script_url = "C:\MAXIS-scripts\notes\caf-testing.vbs"
    Else
        testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/notes/caf-testing.vbs"
    End If
    Call run_from_GitHub(testing_script_url)
End if

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("04/10/2019", "There was a bug that sometimes made the dialogs write over each other and be illegible, updated the script to keep this from happening.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
CALL changelog_update("10/17/2018", "Updated dialog box to reflect currecnt application process.", "MiKayla Handley, Hennepin County")
call changelog_update("05/05/2018", "Added autofill functionality for the DIET panel. NAT errors have been resolved.", "Ilse Ferris, Hennepin County")
call changelog_update("05/04/2018", "Removed autofill functionality for the DIET panel temporarily until MAXIS help desk can resolve NAT errors.", "Ilse Ferris, Hennepin County")
call changelog_update("01/11/2017", "Adding functionality to offer a TIKL for 12 month contact on 24 month SNAP renewals.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.

'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 30, 15, MAXIS_footer_month
  EditBox 110, 25, 30, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "CASH", cash_checkbox
  CheckBox 50, 60, 30, 10, "HC", HC_checkbox
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 135, 60, 35, 10, "EMER", EMER_checkbox
  DropListBox 70, 80, 75, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"Addendum", CAF_type
  ButtonGroup ButtonPressed
    OkButton 35, 100, 50, 15
    CancelButton 95, 100, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 65, 10, "Footer month/year: "
  GroupBox 5, 45, 170, 30, "Programs applied for"
  Text 30, 85, 35, 10, "CAF type:"
EndDialog

'initial dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If CAF_type = "Select One:" then err_msg = err_msg & vbnewline & "* You must select the CAF type."
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You need to type a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call check_for_MAXIS(False)	'checking for an active MAXIS session
MAXIS_footer_month_confirmation	'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
	IF SNAP_checkbox = checked THEN																															'checking for SNAP 24 month renewals.'
		EMWriteScreen "X", 05, 58																																	'opening the FS revw screen.
		transmit
		EMReadScreen SNAP_recert_date, 8, 9, 64
		PF3
		SNAP_recert_date = replace(SNAP_recert_date, " ", "/")																		'replacing the read blank spaces with / to make it a date
		SNAP_recert_compare_date = dateadd("m", "12", MAXIS_footer_month & "/01/" & MAXIS_footer_year)		'making a dummy variable to compare with, by adding 12 months to the requested footer month/year.
		IF datediff("d", SNAP_recert_compare_date, SNAP_recert_date) > 0 THEN											'If the read recert date is more than 0 days away from 12 months plus the MAXIS footer month/year then it is likely a 24 month period.'
			SNAP_recert_is_likely_24_months = TRUE
		ELSE
			SNAP_recert_is_likely_24_months = FALSE																									'otherwise if we don't we set it as false
		END IF
	END IF
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

'I put these sections in here, just because SHEL should come before HEST, it just looks cleaner.
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

'Now it grabs the rest of the info, not dependent on which programs are selected.
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "CASH, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
If CAF_type <> "Recertification" then TIKL_checkbox = checked

'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Do
	Do
		Do
			Do

                BeginDialog Dialog1, 0, 0, 451, 290, "CAF dialog part 1"
                  EditBox 60, 5, 50, 15, CAF_datestamp
                  ComboBox 175, 5, 70, 15, " "+chr(9)+"phone"+chr(9)+"office", interview_type
                  CheckBox 255, 5, 65, 10, "Used Interpreter", Used_Interpreter_checkbox
                  EditBox 60, 25, 50, 15, interview_date
                  ComboBox 230, 25, 95, 15,  " "+chr(9)+"Select One:"+chr(9)+"Fax"+chr(9)+"Mail"+chr(9)+"Office"+chr(9)+"Online", how_app_rcvd
                  ComboBox 220, 45, 105, 15, " "+chr(9)+"DHS-2128 (LTC Renewal)"+chr(9)+"DHS-3417B (Req. to Apply...)"+chr(9)+"DHS-3418 (HC Renewal)"+chr(9)+"DHS-3531 (LTC Application)"+chr(9)+"DHS-3876 (Certain Pops App)"+chr(9)+"DHS-6696(MNsure HC App)", HC_document_received
                  EditBox 390, 45, 50, 15, HC_datestamp
                  EditBox 75, 70, 370, 15, HH_comp
                  EditBox 35, 90, 200, 15, cit_id
                  EditBox 265, 90, 180, 15, IMIG
                  EditBox 60, 110, 120, 15, AREP
                  EditBox 270, 110, 175, 15, SCHL
                  EditBox 60, 130, 210, 15, DISA
                  EditBox 310, 130, 135, 15, FACI
                  EditBox 35, 160, 410, 15, PREG
                  EditBox 35, 180, 410, 15, ABPS
                  EditBox 35, 200, 410, 15, EMPS
                  If worker_county_code = "x127" or worker_county_code = "x162" then CheckBox 35, 220, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
                  EditBox 55, 235, 390, 15, verifs_needed
                  ButtonGroup ButtonPressed
                    PushButton 340, 270, 50, 15, "NEXT", next_to_page_02_button
                    CancelButton 395, 270, 50, 15
                    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                    PushButton 335, 25, 45, 10, "next panel", next_panel_button
                    PushButton 395, 25, 45, 10, "next memb", next_memb_button
                    PushButton 5, 75, 60, 10, "HH comp/EATS:", EATS_button
                    PushButton 240, 95, 20, 10, "IMIG:", IMIG_button
                    PushButton 5, 115, 25, 10, "AREP/", AREP_button
                    PushButton 30, 115, 25, 10, "ALTP:", ALTP_button
                    PushButton 190, 115, 25, 10, "SCHL/", SCHL_button
                    PushButton 215, 115, 25, 10, "STIN/", STIN_button
                    PushButton 240, 115, 25, 10, "STEC:", STEC_button
                    PushButton 5, 135, 25, 10, "DISA/", DISA_button
                    PushButton 30, 135, 25, 10, "PDED:", PDED_button
                    PushButton 280, 135, 25, 10, "FACI:", FACI_button
                    PushButton 5, 165, 25, 10, "PREG:", PREG_button
                    PushButton 5, 185, 25, 10, "ABPS:", ABPS_button
                    PushButton 5, 205, 25, 10, "EMPS", EMPS_button
                    PushButton 10, 270, 20, 10, "DWP", ELIG_DWP_button
                    PushButton 30, 270, 15, 10, "FS", ELIG_FS_button
                    PushButton 45, 270, 15, 10, "GA", ELIG_GA_button
                    PushButton 60, 270, 15, 10, "HC", ELIG_HC_button
                    PushButton 75, 270, 20, 10, "MFIP", ELIG_MFIP_button
                    PushButton 95, 270, 20, 10, "MSA", ELIG_MSA_button
                    PushButton 130, 270, 25, 10, "ADDR", ADDR_button
                    PushButton 155, 270, 25, 10, "MEMB", MEMB_button
                    PushButton 180, 270, 25, 10, "MEMI", MEMI_button
                    PushButton 205, 270, 25, 10, "PROG", PROG_button
                    PushButton 230, 270, 25, 10, "REVW", REVW_button
                    PushButton 255, 270, 25, 10, "SANC", SANC_button
                    PushButton 280, 270, 25, 10, "TIME", TIME_button
                    PushButton 305, 270, 25, 10, "TYPE", TYPE_button
                  Text 335, 50, 55, 10, "HC datestamp:"
                  Text 5, 95, 25, 10, "CIT/ID:"
                  Text 5, 240, 50, 10, "Verifs needed:"
                  GroupBox 5, 260, 115, 25, "ELIG panels:"
                  GroupBox 125, 260, 210, 25, "other STAT panels:"
                  GroupBox 330, 5, 115, 35, "STAT-based navigation"
                  Text 5, 10, 55, 10, "CAF datestamp:"
                  Text 5, 30, 55, 10, "Interview date:"
                  Text 120, 10, 50, 10, "Interview type:"
                  Text 5, 50, 210, 10, "If HC applied for (or recertifying): what document was received?:"
                  Text 120, 30, 110, 10, "How was application received?:"
                EndDialog

				err_msg = ""
				Dialog Dialog1			'Displays the first dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				If CAF_datestamp = "" or len(CAF_datestamp) > 10 THEN err_msg = "Please enter a valid application datestamp."
				If err_msg <> "" THEN Msgbox err_msg
			Loop until ButtonPressed = next_to_page_02_button and err_msg = ""
			Do
				Do
                    BeginDialog Dialog1, 0, 0, 451, 305, "CAF dialog part 2"
                      EditBox 70, 55, 370, 15, earned_income
                      EditBox 80, 75, 360, 15, unearned_income
                      EditBox 115, 95, 325, 15, notes_on_income
                      EditBox 85, 115, 355, 15, income_changes
                      EditBox 160, 135, 280, 15, is_any_work_temporary
                      EditBox 65, 160, 375, 15, notes_on_abawd
                      EditBox 65, 180, 375, 15, SHEL_HEST
                      EditBox 65, 200, 250, 15, COEX_DCEX
                      EditBox 65, 220, 375, 15, CASH_ACCTs
                      EditBox 155, 240, 285, 15, other_assets
                      EditBox 55, 265, 385, 15, verifs_needed
                      ButtonGroup ButtonPressed
                        PushButton 270, 290, 60, 10, "previous page", previous_to_page_01_button
                        PushButton 335, 285, 50, 15, "NEXT", next_to_page_03_button
                        CancelButton 390, 285, 50, 15
                        PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
                        PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
                        PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
                        PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
                        PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
                        PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
                        PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
                        PushButton 150, 15, 25, 10, "BUSI", BUSI_button
                        PushButton 175, 15, 25, 10, "JOBS", JOBS_button
                        PushButton 200, 15, 25, 10, "PBEN", PBEN_button
                        PushButton 225, 15, 25, 10, "RBIC", RBIC_button
                        PushButton 250, 15, 25, 10, "UNEA", UNEA_button
                        PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                        PushButton 335, 25, 45, 10, "next panel", next_panel_button
                        PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                        PushButton 395, 25, 45, 10, "next memb", next_memb_button
                        PushButton 10, 100, 100, 10, "Notes on Income and Budget", income_notes_button
                        PushButton 10, 120, 70, 10, "STWK/inc. changes:", STWK_button
                        PushButton 5, 165, 60, 10, "ABAWD/WREG:", WREG_button
                        PushButton 5, 185, 25, 10, "SHEL/", SHEL_button
                        PushButton 30, 185, 25, 10, "HEST:", HEST_button
                        PushButton 5, 205, 25, 10, "COEX/", COEX_button
                        PushButton 30, 205, 25, 10, "DCEX:", DCEX_button
                        PushButton 5, 225, 25, 10, "CASH/", CASH_button
                        PushButton 30, 225, 30, 10, "ACCTs:", ACCT_button
                        PushButton 5, 245, 25, 10, "CARS/", CARS_button
                        PushButton 30, 245, 25, 10, "REST/", REST_button
                        PushButton 55, 245, 25, 10, "SECU/", SECU_button
                        PushButton 80, 245, 25, 10, "TRAN/", TRAN_button
                        PushButton 105, 245, 45, 10, "other assets:", OTHR_button
                      GroupBox 145, 5, 135, 25, "Income panels"
                      GroupBox 330, 5, 115, 35, "STAT-based navigation"
                      Text 15, 60, 55, 10, "Earned income:"
                      Text 15, 80, 60, 10, "Unearned income:"
                      Text 10, 140, 150, 10, "Is any work temporary? If so, explain details:"
                      Text 5, 270, 50, 10, "Verifs needed:"
                      GroupBox 5, 5, 130, 25, "ELIG panels:"
                      GroupBox 5, 40, 440, 115, "Income info: Please explain budged/excluded income for the case. If income has ended, case note the info and remove old panel(s)."
                    EndDialog

					err_msg = ""
					income_note_error_msg = ""
					Dialog Dialog1			'Displays the second dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = income_notes_button Then
                        BeginDialog Dialog1, 0, 0, 351, 215, "Explanation of Income"
                          CheckBox 10, 30, 325, 10, "JOBS - Income detail on previous note(s)", see_other_note_checkbox
                          CheckBox 10, 45, 325, 10, "JOBS - Income has not been verified and detail will be entered when received.", not_verified_checkbox
                          CheckBox 10, 60, 325, 10, "JOBS - Client has confirmed that JOBS income is expected to continue at this rate and hours.", jobs_anticipated_checkbox
                          CheckBox 10, 75, 330, 10, "JOBS - This is a new job and actual check stubs are not available, advised client that if actual pay", new_jobs_checkbox
                          CheckBox 10, 100, 325, 10, "BUSI - Client has confirmed that BUSI income is expected to continue at this rate and hours.", busi_anticipated_checkbox
                          CheckBox 10, 115, 250, 10, "BUSI - Client has agreed to the self-employment budgeting method used.", busi_method_agree_checkbox
                          CheckBox 10, 130, 325, 10, "RBIC - Client has confirmed that RBIC income is expected to continue at this rate and hours.", rbic_anticipated_checkbox
                          CheckBox 10, 145, 325, 10, "UNEA - Client has confirmed that UNEA income is expected to continue at this rate and hours.", unea_anticipated_checkbox
                          CheckBox 10, 160, 315, 10, "UNEA - Client has applied for unemployment benefits but no determination made at this time.", ui_pending_checkbox
                          CheckBox 45, 170, 225, 10, "Check here to have the script set a TIKL to check UI in two weeks.", tikl_for_ui
                          CheckBox 10, 185, 150, 10, "NONE - This case has no income reported.", no_income_checkbox
                          ButtonGroup ButtonPressed
                            PushButton 240, 195, 50, 15, "Insert", add_to_notes_button
                            CancelButton 295, 195, 50, 15
                          Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
                          Text 45, 85, 315, 10, "varies significantly, client should provide proof of this difference to have benefits adjusted."
                        EndDialog

						Dialog Dialog1
						If ButtonPressed = add_to_notes_button Then
                            If see_other_note_checkbox Then notes_on_income = notes_on_income & "; Full detail about income can be found in previous note(s)."
                            If not_verified_checkbox Then notes_on_income = notes_on_income & "; This income has not been fully verified and information about income for budget will be noted when the verification is received."
							If jobs_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from jobs to continue at this amount."
							If new_jobs_checkbox = checked Then notes_on_income = notes_on_income & "; This is a new job and actual check stubs have not been received, advised client to provide proof once pay is received if the income received differs significantly."
							If busi_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects all income from self employment to continue at this amount."
							If busi_method_agree_checkbox = checked Then notes_on_income = notes_on_income & "; Explained to client the self employment budgeting methods and client agreed to the method used."
							If rbic_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects roomer/boarder income to continue at this amount."
							If unea_anticipated_checkbox = checked Then notes_on_income = notes_on_income & "; Client expects unearned income to continue at this amount."
							If ui_pending_checkbox = checked Then notes_on_income = notes_on_income & "; Client has applied for Unemployment Income recently but request is still pending, will need to be reviewed soon for changes."
							If tikl_for_ui = checked Then notes_on_income = notes_on_income & " TIKL set to request an update on Unemployment Income."
							If no_income_checkbox = checked Then notes_on_income = notes_on_income & "; Client has reported they have no income and do not expect any changes to this at this time."
							If left(notes_on_income, 1) = ";" Then notes_on_income = right(notes_on_income, len(notes_on_income) - 1)
						End If
					End If
					IF (earned_income <> "" AND trim(notes_on_income) = "") OR (unearned_income <> "" AND notes_on_income = "") THEN income_note_error_msg = True
					If err_msg <> "" THEN Msgbox err_msg
				Loop until ButtonPressed = (next_to_page_03_button AND err_msg = "") or (ButtonPressed = previous_to_page_01_button AND err_msg = "")		'If you press either the next or previous button, this loop ends
				If ButtonPressed = previous_to_page_01_button then exit do		'If the button was previous, it exits this do loop and is caught in the next one, which sends you back to Dialog 1 because of the "If ButtonPressed = previous_to_page_01_button then exit do" later on
				Do
                    BeginDialog Dialog1, 0, 0, 451, 405, "CAF dialog part 3"
                      EditBox 60, 45, 385, 15, INSA
                      EditBox 35, 65, 410, 15, ACCI
                      EditBox 35, 85, 175, 15, DIET
                      EditBox 245, 85, 200, 15, BILS
                      EditBox 35, 105, 285, 15, FMED
                      EditBox 390, 105, 55, 15, retro_request
                      EditBox 180, 130, 265, 15, reason_expedited_wasnt_processed
                      EditBox 100, 150, 345, 15, FIAT_reasons
                      CheckBox 15, 190, 80, 10, "Application signed?", application_signed_checkbox
                      CheckBox 15, 205, 65, 10, "Appt letter sent?", appt_letter_sent_checkbox
                      CheckBox 15, 220, 150, 10, "Client willing to participate with E and T", E_and_T_checkbox
                      CheckBox 15, 235, 70, 10, "EBT referral sent?", EBT_referral_checkbox
                      CheckBox 115, 190, 50, 10, "eDRS sent?", eDRS_sent_checkbox
                      CheckBox 115, 205, 50, 10, "Expedited?", expedited_checkbox
                      CheckBox 115, 235, 70, 10, "IAAs/OMB given?", IAA_checkbox
                      CheckBox 200, 190, 115, 10, "Informed client of recert period?", recert_period_checkbox
                      CheckBox 200, 205, 80, 10, "Intake packet given?", intake_packet_checkbox
                      CheckBox 200, 220, 105, 10, "Managed care packet sent?", managed_care_packet_checkbox
                      CheckBox 200, 235, 105, 10, "Managed care referral made?", managed_care_referral_checkbox
                      CheckBox 345, 190, 65, 10, "R/R explained?", R_R_checkbox
                      CheckBox 345, 205, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
                      CheckBox 345, 220, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
                      CheckBox 345, 235, 95, 10, "Workforce referral made?", WF1_checkbox
                      EditBox 55, 260, 230, 15, other_notes
                      EditBox 55, 280, 390, 15, verifs_needed
                      EditBox 55, 300, 390, 15, actions_taken
                      ComboBox 330, 260, 115, 15, " "+chr(9)+"incomplete"+chr(9)+"approved", CAF_status
                      CheckBox 15, 335, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
                      CheckBox 15, 350, 200, 10, "Check here to create a TIKL to deny at the 30/45 day mark.", TIKL_checkbox
                      CheckBox 15, 365, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
                      EditBox 395, 345, 50, 15, worker_signature
                      ButtonGroup ButtonPressed
                        PushButton 290, 370, 45, 10, "prev. page", previous_to_page_02_button
                        OkButton 340, 365, 50, 15
                        CancelButton 395, 365, 50, 15
                        PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
                        PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
                        PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
                        PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
                        PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
                        PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
                        PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
                        PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
                        PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
                        PushButton 335, 25, 45, 10, "next panel", next_panel_button
                        PushButton 395, 25, 45, 10, "next memb", next_memb_button
                      GroupBox 5, 5, 130, 25, "ELIG panels:"
                      GroupBox 330, 5, 115, 35, "STAT-based navigation"
                      ButtonGroup ButtonPressed
                        PushButton 5, 50, 25, 10, "INSA/", INSA_button
                        PushButton 30, 50, 25, 10, "MEDI:", MEDI_button
                        PushButton 5, 70, 25, 10, "ACCI:", ACCI_button
                        PushButton 5, 90, 25, 10, "DIET:", DIET_button
                        PushButton 5, 110, 25, 10, "FMED:", FMED_button
                      Text 5, 135, 170, 10, "Reason expedited wasn't processed (if applicable):"
                      Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
                      GroupBox 5, 175, 440, 75, "Common elements workers should case note:"
                      Text 5, 265, 50, 10, "Other notes:"
                      Text 290, 265, 40, 10, "CAF status:"
                      Text 5, 285, 50, 10, "Verifs needed:"
                      Text 5, 305, 50, 10, "Actions taken:"
                      GroupBox 5, 320, 280, 60, "Actions the script can do:"
                      Text 330, 350, 60, 10, "Worker signature:"
                      ButtonGroup ButtonPressed
                        PushButton 325, 110, 60, 10, "Retro Req. date:", HCRE_button
                        PushButton 215, 90, 25, 10, "BILS:", BILS_button
                    EndDialog

					err_msg = ""
					Dialog Dialog1			'Displays the third dialog
					cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
					MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
					If ButtonPressed = previous_to_page_02_button then exit do		'Exits this do...loop here if you press previous. The second ""loop until ButtonPressed = -1" gets caught, and it loops back to the "Do" after "Loop until ButtonPressed = next_to_page_02_button"
					If actions_taken = "" THEN err_msg = err_msg & vbCr & "Please complete actions taken section."    'creating err_msg if required items are missing
					If income_note_error_msg = True THEN err_msg = err_msg & VbCr & "Income for this case was found in MAXIS. Please complete the 'notes on income and budget' field."
					If worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature."
					If CAF_status = " " THEN err_msg = err_msg & vbCr & "Please select a CAF Status."
					If err_msg <> "" THEN Msgbox err_msg
				Loop until (ButtonPressed = -1 and err_msg = "") or (ButtonPressed = previous_to_page_02_button and err_msg = "")		'If OK or PREV, it exits the loop here, which is weird because the above also causes it to exit
			Loop until ButtonPressed = -1	'Because this is in here a second time, it triggers a return to the "Dialog CAF_dialog_02" line, where all those "DOs" start again!!!!!
			If ButtonPressed = previous_to_page_01_button then exit do 	'This exits this particular loop again for prev button on page 2, which sends you back to page 1!!
		Loop until err_msg = ""		'Loops all of that until those four sections are finished. Let's move that over to those particular pages. Folks would be less angry that way I bet.
		CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
	Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
	call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call check_for_maxis(FALSE)  'allows for looping to check for maxis after worker has complete dialog box so as not to lose a giant CAF case note if they get timed out while writing.

'This code will update the interview date in PROG.
If CAF_type <> "Recertification" AND CAF_type <> "Addendum" Then        'Interview date is not on PROG for recertifications or addendums
    If SNAP_checkbox = checked OR cash_checkbox = checked Then          'Interviews are only required for Cash and SNAP
        intv_date_needed = FALSE
        Call navigate_to_MAXIS_screen("STAT", "PROG")                   'Going to STAT to check to see if there is already an interview indicated.

        If SNAP_checkbox = checked Then                                 'If the script is being run for a SNAP interview
            EMReadScreen entered_intv_date, 8, 10, 55                   'REading what is entered in the SNAP interview
            'MsgBox "SNAP interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE  'If this is blank - the script needs to prompt worker to update it
        End If

        If cash_checkbox = checked THen                             'If the script is bring run for a Cash interview
            EMReadScreen cash_one_app, 8, 6, 33                     'First the script needs to identify if it is cash 1 or cash 2 that has the application information
            EMReadScreen cash_two_app, 8, 7, 33
            EMReadScreen grh_cash_app, 8, 9, 33

            cash_one_app = replace(cash_one_app, " ", "/")          'Turning this in to a date format
            cash_two_app = replace(cash_two_app, " ", "/")
            grh_cash_app = replace(grh_cash_app, " ", "/")

            If cash_one_app <> "__/__/__" Then      'Error handling - VB doesn't like date comparisons with non-dates
                if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6     'If date of application on PROG matches script date of applicaton
            End If
            If cash_two_app <> "__/__/__" Then
                if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
            End If

            If grh_cash_app <> "__/__/__" Then
                if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
            End If

            EMReadScreen entered_intv_date, 8, prog_row, 55                     'Reading the right interview date with row defined above
            'MsgBox "Cash interview date - " & entered_intv_date
            If entered_intv_date = "__ __ __" Then intv_date_needed = TRUE      'If this is blank - script needs to prompt worker to have it updated
        End If

        If intv_date_needed = TRUE Then         'If previous code has determined that PROG needs to be updated
            If SNAP_checkbox = checked Then prog_update_SNAP_checkbox = checked     'Auto checking based on the programs the script is being run for.
            If cash_checkbox = checked Then prog_update_cash_checkbox = checked

            'Dialog code
            BeginDialog Dialog1, 0, 0, 231, 130, "Update PROG?"
              OptionGroup RadioGroup1
                RadioButton 10, 10, 155, 10, "YES! Update PROG with the Interview Date", confirm_update_prog
                RadioButton 10, 60, 90, 10, "No, do not update PROG", do_not_update_prog
              EditBox 165, 5, 50, 15, interview_date
              CheckBox 25, 25, 30, 10, "SNAP", prog_update_SNAP_checkbox
              CheckBox 25, 40, 30, 10, "CASH", prog_update_cash_checkbox
              Text 20, 75, 200, 10, "Reason PROG should not be updated with the Interview Date:"
              EditBox 20, 90, 195, 15, no_update_reason
              ButtonGroup ButtonPressed
                OkButton 175, 110, 50, 15
            EndDialog

            'Running the dialog
            Do
                err_msg = ""
                Dialog Dialog1
                'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
                If do_not_update_prog = 1 AND no_update_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
                IF confirm_update_prog = 1 AND prog_update_SNAP_checkbox = unchecked AND prog_update_cash_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select either CASH or SNAP to have updated on PROG."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop until err_msg = ""

            If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
                CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

                keep_footer_month = MAXIS_footer_month      'Saving the footer month and year that was determined earlier in the script. It needs t obe changed for nav functions to work correctly
                keep_footer_year = MAXIS_footer_year

                app_month = DatePart("m", CAF_datestamp)    'Setting the footer month and year to the app month.
                app_year = DatePart("yyyy", CAF_datestamp)

                MAXIS_footer_month = right("00" & app_month, 2)
                MAXIS_footer_year = right(app_year, 2)

                CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
                PF9                                             'Edit

                intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
                intv_day = DatePart("d", interview_date)
                intv_yr = DatePart("yyyy", interview_date)

                intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
                intv_day = right("00"&intv_day, 2)
                intv_yr = right(intv_yr, 2)

                If prog_update_SNAP_checkbox = checked Then     'If it was selected to SNAP interview to be updated
                    programs_w_interview = "SNAP"               'Setting a variable for case noting

                    EMWriteScreen intv_mo, 10, 55               'SNAP is easy because there is only one area for interview - the variables go there
                    EMWriteScreen intv_day, 10, 58
                    EMWriteScreen intv_yr, 10, 61
                End If

                If prog_update_cash_checkbox = checked Then     'If it was selected to update for Cash
                    If programs_w_interview = "" Then programs_w_interview = "CASH"     'variable for the case note
                    If programs_w_interview <> "" Then programs_w_interview = "SNAP and CASH"
                    EMReadScreen cash_one_app, 8, 6, 33     'Reading app dates of both cash lines
                    EMReadScreen cash_two_app, 8, 7, 33
                    EMReadScreen grh_cash_app, 8, 9, 33

                    cash_one_app = replace(cash_one_app, " ", "/")      'Formatting as dates
                    cash_two_app = replace(cash_two_app, " ", "/")
                    grh_cash_app = replace(grh_cash_app, " ", "/")

                    If cash_one_app <> "__/__/__" Then              'Comparing them to the date of application to determine which row to use
                        if DateDiff("d", cash_one_app, CAF_datestamp) = 0 then prog_row = 6
                    End If
                    If cash_two_app <> "__/__/__" Then
                        if DateDiff("d", cash_two_app, CAF_datestamp) = 0 then prog_row = 7
                    End If

                    If grh_cash_app <> "__/__/__" Then
                        if DateDiff("d", grh_cash_app, CAF_datestamp) = 0 then prog_row = 9
                    End If

                    EMWriteScreen intv_mo, prog_row, 55     'Writing the interview date in
                    EMWriteScreen intv_day, prog_row, 58
                    EMWriteScreen intv_yr, prog_row, 61
                End If

                transmit                                    'Saving the panel

                MAXIS_footer_month = keep_footer_month      'resetting the footer month and year so the rest of the script uses the worker identified footer month and year.
                MAXIS_footer_year = keep_footer_year
            End If
        ENd If
    End If
End If
'MsgBox "PROG Stuff done"
'MsgBox confirm_update_prog

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and CAF_type <> "Recertification" then
	call navigate_to_MAXIS_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
		EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
		If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
		EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
		If PND2_HC_status_check = "P" then
			EMWriteScreen "x", PND2_row, 3
			transmit
			person_delay_row = 7
			Do
				EMReadScreen person_delay_check, 1, person_delay_row, 39
				If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
				person_delay_row = person_delay_row + 2
			Loop until person_delay_check = " " or person_delay_row > 20
			PF3
		End if
		EMReadScreen additional_app_check, 14, PND2_row + 1, 17
		If additional_app_check <> "ADDITIONAL APP" then exit for
		PND2_row = PND2_row + 1
	next
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
	End if
End if

'Going to TIKL. Now using the write TIKL function
If TIKL_checkbox = checked and CAF_type <> "Recertification" then
	If cash_checkbox = checked or EMER_checkbox = checked or SNAP_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 30 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 30 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 30, 5, 18)
			If cash_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "Cash/"
			If SNAP_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "SNAP/"
			If EMER_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "EMER/"
			TIKL_msg_one = Left(TIKL_msg_one, (len(TIKL_msg_one) - 1))
			TIKL_msg_one = TIKL_msg_one & " has been pending for 30 days. Evaluate for possible denial."
			Call write_variable_in_TIKL (TIKL_msg_one)
			PF3
		End If
	End if
	If HC_checkbox = checked then
		If DateDiff ("d", CAF_datestamp, date) > 45 Then 'Error handling to prevent script from attempting to write a TIKL in the past
			MsgBox "Cannot set TIKL as CAF Date is over 45 days old and TIKL would be in the past. You must manually track."
		Else
			call navigate_to_MAXIS_screen("dail", "writ")
			call create_MAXIS_friendly_date(CAF_datestamp, 45, 5, 18)
			Call write_variable_in_TIKL ("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
			PF3
		End If
	End if
End if
If client_delay_TIKL_checkbox = checked then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	Call write_variable_in_TIKL (">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<")
	PF3
End if

IF tikl_for_ui THEN
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	two_weeks_from_now = DateAdd("d", 14, date)
	call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
	call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
	PF3
END IF
'--------------------END OF TIKL BUSINESS

'Navigates to case note, and checks to make sure we aren't in inquiry.
Call start_a_blank_CASE_NOTE

'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***" & CAF_type & CAF_status & "***")
IF move_verifs_needed = TRUE THEN
	CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
	CALL write_variable_in_CASE_NOTE("------------------------------")
End if
CALL write_bullet_and_variable_in_CASE_NOTE("CAF datestamp", CAF_datestamp)
If Used_Interpreter_checkbox = checked then
	CALL write_variable_in_CASE_NOTE("* Interview type: " & interview_type & " w/ interpreter")
Else
	CALL write_bullet_and_variable_in_CASE_NOTE("Interview type", interview_type)
End if
CALL write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
'If intv_date_needed = FALSE THen CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG by worker prior to script run.")
If confirm_update_prog = 1 Then CALL write_variable_in_CASE_NOTE("* Interview date entered on PROG for " & programs_w_interview)
If do_not_update_prog = 1 Then CALL write_bullet_and_variable_in_CASE_NOTE("PROG WAS NOT UPDATED WITH INTERVIEW DATE, because", no_update_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("HC document received", HC_document_received)
CALL write_bullet_and_variable_in_CASE_NOTE("HC datestamp", HC_datestamp)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("How CAF was received", how_app_rcvd)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
CALL write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
CALL write_bullet_and_variable_in_CASE_NOTE("SCHL/STIN/STEC", SCHL)
CALL write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
CALL write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
Call write_bullet_and_variable_in_CASE_NOTE("EMPS", EMPS)
If MFIP_DVD_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MFIP financial orientation DVD sent to participant(s).")
CALL write_bullet_and_variable_in_CASE_NOTE("ABPS", ABPS)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned inc.", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes on income and budget", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("STWK/inc. changes", income_changes)
CALL write_bullet_and_variable_in_CASE_NOTE("ABAWD Notes/WREG", notes_on_abawd)
CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary", is_any_work_temporary)
CALL write_bullet_and_variable_in_CASE_NOTE("SHEL/HEST", SHEL_HEST)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("CASH/ACCTs", CASH_ACCTs)
CALL write_bullet_and_variable_in_CASE_NOTE("Other assets", other_assets)
CALL write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
CALL write_bullet_and_variable_in_CASE_NOTE("ACCI", ACCI)
CALL write_bullet_and_variable_in_CASE_NOTE("DIET", DIET)
CALL write_bullet_and_variable_in_CASE_NOTE("BILS", BILS)
CALL write_bullet_and_variable_in_CASE_NOTE("FMED", FMED)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Request (IF applicable)", retro_request)
IF application_signed_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF appt_letter_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Appointment letter was sent before interview.")
IF EBT_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* EBT referral made for client.")
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
IF managed_care_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received managed care packet.")
IF managed_care_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Managed care referral made.")
IF R_R_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* R/R explained to client.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF E_and_T_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client is willing to participate with E&T.")
IF recert_period_checkbox = checked THEN call write_variable_in_CASE_NOTE("* Informed client of recert period.")
IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
CALL write_bullet_and_variable_in_CASE_NOTE("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_recert_is_likely_24_months = TRUE THEN					'if we determined on stat/revw that the next SNAP recert date isn't 12 months beyond the entered footer month/year
	TIKL_for_24_month = msgbox("Your SNAP recertification date is listed as " & SNAP_recert_date & " on STAT/REVW. Do you want set a TIKL on " & dateadd("m", "-1", SNAP_recert_compare_date) & " for 12 month contact?" & vbCR & vbCR & "NOTE: Clicking yes will navigate away from CASE/NOTE saving your case note.", VBYesNo)
	IF TIKL_for_24_month = vbYes THEN 												'if the select YES then we TIKL using our custom functions.
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		CALL create_MAXIS_friendly_date(dateadd("m", "-1", SNAP_recert_compare_date), 0, 5, 18)
		CALL write_variable_in_TIKL("If SNAP is open, review to see if 12 month contact letter is needed. DAIL scrubber can send 12 Month Contact Letter if used on this TIKL.")
	END IF
END IF

end_msg = "Success! CAF has been successfully noted. Please remember to run the Approved Programs, Closed Programs, or Denied Programs scripts if  results have been APP'd."
If do_not_update_prog = 1 Then end_msg = end_msg & vbNewLine & vbNewLine & "It was selected that PROG would NOT be updated because " & no_update_reason
script_end_procedure(end_msg)
