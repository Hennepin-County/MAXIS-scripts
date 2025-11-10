'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CSR.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600          'manual run time in seconds
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
call changelog_update("11/10/2025", "Improved background script functionality, streamlined dialog options, added signature details to final dialog, and removed unneeded dialog fields.", "Mark Riegel, Hennepin County")
call changelog_update("09/23/2025", "Returned the Health Care Programs option for CSR Processing.##~##(It had previously been removed during the PHE.)", "Casey Love, Hennepin County")
call changelog_update("07/11/2025", "Added MFIP and GA to CSR Program Selection to support new CSR processing on these programs.", "Casey Love, Hennepin County")
call changelog_update("06/01/2022", "Removed Paperless IR and Health Care Programs selections during Public Health Emergency.", "Ilse Ferris, Hennepin County") '#863
call changelog_update("05/26/2022", "Fixed bug that did not recognize CSR status as mandatory. Made background stability updates.", "Ilse Ferris, Hennepin County") '#863
call changelog_update("05/21/2021", "Updated browser to default when opening SIR from Internet Explorer to Edge.", "Ilse Ferris")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("12/21/2019", "Updated the script to carry the Footer Month and Year to the MA Approval case note when 'Processing Paperless IR' is checked for an LTC case.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
call changelog_update("12/22/2018", "Added closing message reminder about accepting all ECF work items for CSR's at the time of processing.", "Ilse Ferris, Hennepin County")
call changelog_update("12/07/2018", "Added Paperless (*) IR Option back, with updated functionality.", "Casey Love, Hennepin County")
call changelog_update("11/27/2018", "Removed Paperless (*) IR Option as this CASE/NOTE was insufficient.", "Casey Love, Hennepin County")
call changelog_update("01/17/2017", "This script has been updated to clean up the case note. The script was case noting the ''Verifs Needed'' section twice. This has been resolved.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DEFINING CONSTANTS, ARRAY and BUTTONS===========================================================================

'Buttons Defined
'--Navigation buttons
SHEL_button         = 201
HEST_button         = 202
COEX_button         = 203
DCEX_button         = 204
next_button         = 205
prev_panel_button   = 206
prev_memb_button    = 207
next_panel_button   = 208
next_memb_button    = 209
ELIG_FS_button      = 210
ELIG_HC_button      = 211
ELIG_GRH_button     = 212
ACCT_button         = 213
BUSI_button         = 214
CARS_button         = 215
CASH_button         = 216
JOBS_button         = 217
OTHR_button         = 218
REST_button         = 219
UNEA_button         = 220
SECU_button         = 221
TRAN_button         = 222
MEMB_button         = 223
MEMI_button         = 224
REVW_button         = 225
previous_button     = 226

'--Other buttons
add_to_notes_button       = 227
cancel_income_expl_button = 228
SIR_mail_button           = 229
income_notes_button       = 230

'Defining variables
dialog_count = ""

'DEFINING FUNCTIONS===========================================================================
function button_movement() 	'Dialog movement handling for buttons displayed on the individual form dialogs.
  If dialog_count = 1 Then
    If err_msg = "" Then
      If ButtonPressed = -1 or ButtonPressed = next_button Then dialog_count = 3
      If ButtonPressed = income_notes_button Then dialog_count = 2
    End If
  ElseIf dialog_count = 2 Then
    If ButtonPressed = add_to_notes_button or ButtonPressed = -1 Then
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
      dialog_count = 1
    End If
    If ButtonPressed = cancel_income_expl_button Then dialog_count = 1
  ElseIf dialog_count = 3 Then
    If err_msg = "" Then
      If ButtonPressed = SIR_mail_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
      If ButtonPressed = previous_button Then dialog_count = 1
      If ButtonPressed = -1 Then exit_loop = True
    End If
  End If
end function

'Dialog 1 - CSR First dialog
function CSR_details_dialog()
  dialog_count = 1
  BeginDialog Dialog1, 0, 0, 416, 185, "CSR"
    Text 5, 10, 55, 10, "CSR datestamp:"
    EditBox 70, 5, 50, 15, CSR_datestamp
    Text 5, 30, 40, 10, "CSR status:"
    DropListBox 70, 25, 75, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete", CSR_status
    Text 5, 45, 35, 10, "HH comp:"
    EditBox 70, 40, 220, 15, HH_comp
    Text 5, 60, 55, 10, "Earned income:"
    EditBox 70, 55, 220, 15, earned_income
    Text 5, 75, 60, 10, "Unearned income:"
    EditBox 70, 70, 220, 15, unearned_income
    ButtonGroup ButtonPressed
      PushButton 5, 85, 60, 15, "Notes on Income:", income_notes_button
    EditBox 70, 85, 220, 15, notes_on_income
    Text 5, 105, 60, 10, "Notes on WREG:"
    EditBox 70, 100, 220, 15, notes_on_abawd
    Text 5, 120, 30, 10, "Assets:"
    EditBox 70, 115, 220, 15, assets
    ButtonGroup ButtonPressed
      PushButton 5, 135, 25, 10, "SHEL", SHEL_button
    Text 30, 135, 5, 10, "/"
    ButtonGroup ButtonPressed
      PushButton 35, 135, 25, 10, "HEST:", HEST_button
    EditBox 70, 130, 95, 15, SHEL_HEST
    ButtonGroup ButtonPressed
      PushButton 5, 150, 25, 10, "COEX", COEX_button
    Text 30, 150, 5, 10, "/"
    ButtonGroup ButtonPressed
      PushButton 35, 150, 25, 10, "DCEX:", DCEX_button
    EditBox 70, 145, 95, 15, COEX_DCEX
    ButtonGroup ButtonPressed
      PushButton 360, 165, 50, 15, "Next", next_button
      CancelButton 305, 165, 50, 15
      GroupBox 300, 5, 90, 25, "ELIG panels:"
      ButtonGroup ButtonPressed
      PushButton 305, 15, 25, 10, "FS", ELIG_FS_button
      PushButton 330, 15, 25, 10, "HC", ELIG_HC_button
      PushButton 355, 15, 25, 10, "GRH", ELIG_GRH_button
    GroupBox 300, 35, 90, 55, "Income and asset panels"
      ButtonGroup ButtonPressed
      PushButton 305, 45, 25, 10, "ACCT", ACCT_button
      PushButton 330, 45, 25, 10, "BUSI", BUSI_button
      PushButton 355, 45, 25, 10, "CARS", CARS_button
      PushButton 305, 55, 25, 10, "CASH", CASH_button
      PushButton 330, 55, 25, 10, "JOBS", JOBS_button
      PushButton 355, 55, 25, 10, "OTHR", OTHR_button
      PushButton 305, 65, 25, 10, "REST", REST_button
      PushButton 305, 75, 25, 10, "UNEA", UNEA_button
      PushButton 330, 65, 25, 10, "SECU", SECU_button
      PushButton 355, 65, 25, 10, "TRAN", TRAN_button
    GroupBox 300, 95, 90, 25, "Other STAT panels:"
      ButtonGroup ButtonPressed
      PushButton 305, 105, 25, 10, "MEMB", MEMB_button
      PushButton 330, 105, 25, 10, "MEMI", MEMI_button
      PushButton 355, 105, 25, 10, "REVW", REVW_button
  EndDialog
end function
Dim CSR_datestamp, CSR_status, HH_comp, earned_income, unearned_income, income_notes_button, notes_on_income, notes_on_abawd, assets, SHEL_HEST, COEX_DCEX

'Dialog 2 - CSR Income Notes Dialog
function CSR_income_notes_dialog()
  dialog_count = 2
  BeginDialog Dialog1, 0, 0, 351, 215, "Explanation of Income"
    Text 5, 10, 180, 10, "Check as many explanations of income that apply to this case."
    CheckBox 10, 30, 325, 10, "JOBS - Income detail on previous note(s)", see_other_note_checkbox
    CheckBox 10, 45, 325, 10, "JOBS - Income has not been verified and detail will be entered when received.", not_verified_checkbox
    CheckBox 10, 60, 325, 10, "JOBS - Client has confirmed that JOBS income is expected to continue at this rate and hours.", jobs_anticipated_checkbox
    CheckBox 10, 75, 330, 10, "JOBS - This is a new job and actual check stubs are not available, advised client that if actual pay", new_jobs_checkbox
    Text 45, 85, 315, 10, "varies significantly, client should provide proof of this difference to have benefits adjusted."
    CheckBox 10, 100, 325, 10, "BUSI - Client has confirmed that BUSI income is expected to continue at this rate and hours.", busi_anticipated_checkbox
    CheckBox 10, 115, 250, 10, "BUSI - Client has agreed to the self-employment budgeting method used.", busi_method_agree_checkbox
    CheckBox 10, 130, 325, 10, "RBIC - Client has confirmed that RBIC income is expected to continue at this rate and hours.", rbic_anticipated_checkbox
    CheckBox 10, 145, 325, 10, "UNEA - Client has confirmed that UNEA income is expected to continue at this rate and hours.", unea_anticipated_checkbox
    CheckBox 10, 160, 315, 10, "UNEA - Client has applied for unemployment benefits but no determination made at this time.", ui_pending_checkbox
    CheckBox 45, 170, 225, 10, "Check here to have the script set a TIKL to check UI in two weeks.", tikl_for_ui
    CheckBox 10, 185, 150, 10, "NONE - This case has no income reported.", no_income_checkbox
    ButtonGroup ButtonPressed
      PushButton 295, 195, 50, 15, "Insert", add_to_notes_button
      PushButton 240, 195, 50, 15, "Cancel", cancel_income_expl_button
  EndDialog
end function
Dim see_other_note_checkbox, not_verified_checkbox, jobs_anticipated_checkbox, new_jobs_checkbox, busi_anticipated_checkbox, busi_method_agree_checkbox, rbic_anticipated_checkbox, unea_anticipated_checkbox, ui_pending_checkbox, tikl_for_ui, no_income_checkbox

'Dialog 3 - CSR Details Cont'd Dialog
function CSR_details_cont_dialog()
  dialog_count = 3
  BeginDialog Dialog1, 0, 0, 396, 180, "CSR (cont)"
    Text 5, 10, 50, 10, "FIAT reasons:"
    EditBox 60, 5, 150, 15, FIAT_reasons
    Text 215, 10, 45, 10, "(if applicable)"
    Text 5, 25, 40, 10, "Other notes:"
    EditBox 60, 20, 230, 15, other_notes
    Text 5, 40, 35, 10, "Changes?:"
    EditBox 60, 35, 230, 15, changes
    Text 5, 55, 50, 10, "Verifs needed:"
    EditBox 60, 50, 230, 15, verifs_needed
    Text 5, 70, 50, 10, "Actions taken:"
    EditBox 60, 65, 230, 15, actions_taken
    Text 5, 90, 90, 10, "Signature of Primary Adult:"
    If SNAP_checkbox = 1 Then 
      ComboBox 100, 85, 75, 15, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+"Not Required", signature
    Else
      ComboBox 100, 85, 75, 15, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Blank"+chr(9)+"Not Required", signature
    End If
    Text 180, 90, 30, 10, "Person:"
    DropListBox 210, 85, 80, 15, HH_Memb_DropDown, signature_memb
    GroupBox 5, 105, 280, 30, "If MA-EPD..."
      Text 10, 120, 50, 10, "New premium:"
      EditBox 65, 115, 80, 15, MAEPD_premium
      CheckBox 155, 120, 65, 10, "Emailed MADE?", MADE_checkbox
      PushButton 225, 115, 50, 15, "SIR mail", SIR_mail_button
      CheckBox 5, 140, 110, 10, "Send forms to AREP?", sent_arep_checkbox
    ButtonGroup ButtonPressed
      OkButton 340, 160, 50, 15
      CancelButton 290, 160, 50, 15
      PushButton 5, 160, 60, 15, "Previous", previous_button
    GroupBox 300, 5, 90, 25, "ELIG panels:"
      ButtonGroup ButtonPressed
      PushButton 305, 15, 25, 10, "FS", ELIG_FS_button
      PushButton 330, 15, 25, 10, "HC", ELIG_HC_button
      PushButton 355, 15, 25, 10, "GRH", ELIG_GRH_button
    GroupBox 300, 35, 90, 55, "Income and asset panels"
      ButtonGroup ButtonPressed
      PushButton 305, 45, 25, 10, "ACCT", ACCT_button
      PushButton 330, 45, 25, 10, "BUSI", BUSI_button
      PushButton 355, 45, 25, 10, "CARS", CARS_button
      PushButton 305, 55, 25, 10, "CASH", CASH_button
      PushButton 330, 55, 25, 10, "JOBS", JOBS_button
      PushButton 355, 55, 25, 10, "OTHR", OTHR_button
      PushButton 305, 65, 25, 10, "REST", REST_button
      PushButton 305, 75, 25, 10, "UNEA", UNEA_button
      PushButton 330, 65, 25, 10, "SECU", SECU_button
      PushButton 355, 65, 25, 10, "TRAN", TRAN_button
    GroupBox 300, 95, 90, 25, "Other STAT panels:"
      ButtonGroup ButtonPressed
      PushButton 305, 105, 25, 10, "MEMB", MEMB_button
      PushButton 330, 105, 25, 10, "MEMI", MEMI_button
      PushButton 355, 105, 25, 10, "REVW", REVW_button
    EndDialog
end function
Dim FIAT_reasons, other_notes, changes, verifs_needed, actions_taken, signature, HH_Memb_DropDown, signature_memb, sent_arep_checkbox, MAEPD_premium, MADE_checkbox

Function dialog_selection(dialog_selected) 	
  'Selects the correct dialog based
  If dialog_selected = 1 then call CSR_details_dialog()
  If dialog_selected = 2 then call CSR_income_notes_dialog()
  If dialog_selected = 3 then call CSR_details_cont_dialog()
End Function

Function dialog_specific_error_handling() 
  If ButtonPressed = next_button or ButtonPressed = previous_button Or ButtonPressed = -1 Then

    If dialog_count = 1 then 
      If isdate(CSR_datestamp) = False THEN err_msg = err_msg & vbCr & "* Please enter the date the CSR was received."
      If CSR_status = "Select one..." THEN err_msg = err_msg & vbCr & "* Please select the status of the CSR."
      If trim(HH_comp) = "" THEN err_msg = err_msg & vbCr & "* Please enter household composition information."
      If (earned_income <> "" AND notes_on_income = "") OR (unearned_income <> "" AND notes_on_income = "") THEN err_msg = err_msg & vbCr & "* You must provide some information about income. Please complete the 'Notes on Income' field."
    End If

    If dialog_count = 3 then 
      If trim(actions_taken) = "" THEN err_msg = err_msg & vbCr & "* Please indicate the actions you have taken."
      If signature = "Select or Type" Then err_msg = err_msg & vbCr & "* You must select or type in one of the signature options."
      If signature_memb = "Select One:" Then err_msg = err_msg & vbCr & "* You must select the household member."
    End If

    If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
  End If
End Function

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
Call check_for_MAXIS(False)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 241, 145, "CSR Evaluation Case Number Dialog"
  Text 5, 5, 225, 10, "Script  purpose: Generates a CASE/NOTE when processing a CSR."
  GroupBox 5, 20, 160, 30, "Programs recertifying"
  CheckBox 10, 35, 30, 10, "SNAP", SNAP_checkbox
  CheckBox 45, 35, 30, 10, "GRH", GRH_checkbox
  CheckBox 75, 35, 30, 10, "MFIP", MFIP_checkbox
  CheckBox 105, 35, 25, 10, "GA", GA_checkbox
  CheckBox 130, 35, 25, 10, "HC", HC_checkbox
  Text 5, 60, 45, 10, "Case number:"
  EditBox 70, 55, 65, 15, MAXIS_case_number
  Text 5, 80, 65, 10, "Footer month/year:"
  EditBox 70, 75, 20, 15, MAXIS_footer_month
  EditBox 95, 75, 20, 15, MAXIS_footer_year
  Text 10, 110, 60, 10, "Worker Signature"
  EditBox 70, 105, 165, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 125, 50, 15
    CancelButton 185, 125, 50, 15
    PushButton 170, 25, 65, 15, "Instructions", script_instructions_btn
EndDialog


'Showing the case number dialog
Do
	DO
    err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
    Call validate_MAXIS_case_number(err_msg, "*")
    Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")

		If SNAP_checkbox + GRH_checkbox + MFIP_checkbox + GA_checkbox + HC_checkbox = 0 Then err_msg = err_msg & vbCr & "* Select at least one program."
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."

		If ButtonPressed = script_instructions_btn Then
			script_instr_url = "https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CSR.docx?"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & script_instr_url
			err_msg = "LOOP"
		End If

		IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE***" & vbCr & err_msg & vbCr & vbCr & "Resolve the following items for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'confirms that footer month/year from dialog matches footer month/year on MAXIS
Call MAXIS_footer_month_confirmation
'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_checkbox = 1 then
  run_from_DAIL = FALSE
  call run_from_GitHub(script_repository &  "dail/paperless-dail.vbs")
End If

'Navigating to STAT/REVW, checking for error prone cases
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged. The script will now end.")

'Creating a custom dialog for determining who the HH members are
Call Generate_Client_List(HH_Memb_DropDown, "Select One:")
call HH_member_custom_dialog(HH_member_array)

'Grabbing SHEL/HEST first, and putting them in this special order that everyone seems to like
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)
'Autofilling assets
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CSR_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Gather phone numbers from ADDR
Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, reservation_name, mail_line_one, mail_line_two, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, address_change_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_one_type, phone_two_type, phone_three_type, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

phone_droplist = "Select or Type"
If phone_one_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_one_number
If phone_two_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_two_number
If phone_three_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_three_number
phone_droplist = phone_droplist+chr(9)+phone_number_selection

'-----------------Creating text for case note
'Programs recertifying case noting info into variable
If MFIP_checkbox = 1 Then programs_recertifying = programs_recertifying & "MFIP, "
If GRH_checkbox = 1 Then programs_recertifying = programs_recertifying & "GRH, "
If GA_checkbox = 1 Then programs_recertifying = programs_recertifying & "GA, "
If SNAP_checkbox = 1 Then programs_recertifying = programs_recertifying & "SNAP, "
If HC_checkbox = 1 Then programs_recertifying = programs_recertifying & "HC, "

programs_recertifying = trim(programs_recertifying)
if right(programs_recertifying, 1) = "," then programs_recertifying = left(programs_recertifying, len(programs_recertifying) - 1)

'Determining the CSR month for header
CSR_month = MAXIS_footer_month & "/" & MAXIS_footer_year

'-------------------------------------------------------------------------------------------------DIALOG

'Start at the first dialog
dialog_count = 1
exit_loop = ""

Do
  Do
    Do
      Dialog1 = "" 'Blanking out previous dialog detail
      Call MAXIS_dialog_navigation
      Call dialog_selection(dialog_count)

      'Blank out variables on each new dialog
      err_msg = ""

      dialog Dialog1 					'Calling a dialog without an assigned variable will call the most recently defined dialog
      cancel_without_confirmation
      Call dialog_specific_error_handling()	'function for error handling of main dialog of forms
      Call button_movement()				'function to move throughout the dialogs
    Loop until err_msg = ""
  Loop until exit_loop = True
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If signature = "Accepted Verbally" Then
  Dialog1 = ""
  BeginDialog Dialog1, 0, 0, 246, 200, "Verbal Signature Record"
    Text 10, 10, 115, 10, "Verbal Signature Accepted for:"
    Text 20, 20, 185, 10, "MEMB " & signature_memb
    Text 20, 50, 190, 20, "To record a verbal signature the date, time and resident phone number needs to be recorded. "
    Text 20, 75, 105, 10, "Signature was accepted at:"
    Text 25, 95, 20, 10, "Date: "
    EditBox 50, 90, 50, 15, verbal_sig_date
    Text 25, 115, 20, 10, "Time: "
    EditBox 50, 110, 50, 15, verbal_sig_time
    Text 20, 140, 85, 10, "Resident Phone Number:"
    DropListBox 110, 135, 95, 45, phone_droplist, verbal_sig_phone_number
    ButtonGroup ButtonPressed
      OkButton 190, 180, 50, 15
    Text 10, 160, 220, 30, "Based on POLI/TEMP 02.05.25 all information here is needed to document the verbal signature. Details will be entered in CASE/NOTE and the WIF in ECF. "
  EndDialog

  Do
    err_msg = ""
    dialog Dialog1
    cancel_without_confirmation

    If IsDate(verbal_sig_date) = False Then err_msg = err_msg & vbCr & "* Enter the date you accepted the verbal signature."
    If IsDate(verbal_sig_time) = True Then
      verbal_sig_time = FormatDateTime(verbal_sig_time, 3)
      If InStr(verbal_sig_time, ":") = 0 Then err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
      verbal_sig_time = replace(verbal_sig_time, ":00 ", " ")
    Else
      err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
    End If
    If verbal_sig_phone_number = "" or verbal_sig_phone_number = "Select or Type" Then err_msg = err_msg & vbCr & "* Phone number detail is required."

    If err_msg <> "" Then MsgBox "*****     NOTICE     *****" & vbCr & "Please resolve to continue:" & vbCr & err_msg
  Loop until err_msg = ""
End If

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF tikl_for_ui THEN Call create_TIKL("Review client's application for Unemployment and request an update if needed.", 14, date, False, TIKL_note_text)

'Writing the case note to MAXIS----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
call write_variable_in_case_note("***" & CSR_month & " CSR received " & CSR_datestamp & ": " & CSR_status & "***")
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
IF move_verifs_needed = TRUE THEN CALL write_variable_in_case_note("---")                               	                'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll add a line separator.
call write_bullet_and_variable_in_case_note("Programs recertifying", programs_recertifying)
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
call write_bullet_and_variable_in_case_note("Notes on Income", notes_on_income)
call write_bullet_and_variable_in_case_note("ABAWD Notes", notes_on_abawd)
call write_bullet_and_variable_in_case_note("Assets", assets)
call write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL_HEST)
call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Changes", changes)
If signature = "Accepted Verbally" Then
  CALL write_variable_in_CASE_NOTE("* * Verbal Signature Accepted:")
  CALL write_variable_in_CASE_NOTE("    - MEMB " & signature_memb)
  CALL write_variable_in_CASE_NOTE("    Signature accepted on " & verbal_sig_date & " at " & verbal_sig_time & ".")
  CALL write_variable_in_CASE_NOTE("    Resident Phone Number: " & verbal_sig_phone_number)
Else
  CALL write_variable_in_CASE_NOTE("* * Signature:" & signature)
  CALL write_variable_in_CASE_NOTE("    - MEMB " & signature_memb)
End If
If sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_bullet_and_variable_in_case_note("MA-EPD premium", MAEPD_premium)
If MADE_checkbox = checked then call write_variable_in_case_note("* Emailed MADE through DHS-SIR.")
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
If tikl_for_ui = 1 then call write_variable_in_CASE_NOTE(TIKL_note_text)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

If paperless_checkbox = unchecked then
    script_end_procedure("Please make sure to accept the Work items in ECF associated with this CSR. Thank you!")
else
    script_end_procedure("")
End if

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/26/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/26/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/26/2022
'--All variables in dialog match mandatory fields-------------------------------05/26/2022
'Review dialog names for content and content fit in dialog----------------------07/11/2025
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------07/11/2025
'--Create a button to reference instructions------------------------------------07/11/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------07/11/2025
'--CASE:NOTE Header doesn't look funky------------------------------------------05/26/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/26/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------07/11/2025
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/26/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------05/26/2022
'--PRIV Case handling reviewed -------------------------------------------------05/26/2022
'--Out-of-County handling reviewed----------------------------------------------05/26/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/26/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/26/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/26/2022------------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/26/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/26/2022
'--Denomination reviewed -------------------------------------------------------05/26/2022
'--Script name reviewed---------------------------------------------------------05/26/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/26/2022------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------07/11/2025
'--comment Code-----------------------------------------------------------------07/11/2025
'--Update Changelog for release/update------------------------------------------07/11/2025
'--Remove testing message boxes-------------------------------------------------07/11/2025
'--Remove testing code/unnecessary code-----------------------------------------07/11/2025
'--Review/update SharePoint instructions----------------------------------------07/11/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/26/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/26/2022
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/26/2022
'--Complete misc. documentation (if applicable)---------------------------------05/26/2022
'--Update project team/issue contact (if applicable)----------------------------05/26/2022
