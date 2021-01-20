'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DOCUMENTS RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation= ucase(windows_user_ID)
name_for_validation = ""
'
' If user_ID_for_validation = "CALO001" Then name_for_validation = "Casey"
' If user_ID_for_validation = "ILFE001" Then name_for_validation = "Ilse"
' If user_ID_for_validation = "WFS395" Then name_for_validation = "MiKayla"
' If user_ID_for_validation = "WFQ898" Then name_for_validation = "Hannah"
' If user_ID_for_validation = "WFK093" Then name_for_validation = "Jessica"
' If user_ID_for_validation = "WFM207" Then name_for_validation = "Mandora"
' If user_ID_for_validation = "WFP803" Then name_for_validation = "Melissa"
' If user_ID_for_validation = "WFC041" Then name_for_validation = "Kerry"
' If user_ID_for_validation = "AAGA001" Then name_for_validation = "Aaron"
' If user_ID_for_validation = "WFJ454" Then name_for_validation = "True"
' If user_ID_for_validation = "WFC719" Then name_for_validation = "Kristen"
' If user_ID_for_validation = "WFE269" Then name_for_validation = "Carrie"
' If user_ID_for_validation = "WFW682" Then name_for_validation = "Osman"
' If user_ID_for_validation = "WFC804" Then name_for_validation = "Shanna"
' If user_ID_for_validation = "WFA168" Then name_for_validation = "Michelle"

If name_for_validation <> "" Then
    MsgBox "Hello " & name_for_validation &  ", you have been selected to test the script NOTES - Documents Received."  & vbNewLine & vbNewLine & "A testing version of the script will now run.  Thank you for taking your time to review our new scripts and functionality as we strive for Continuous Improvement." & vbNewLine & vbNewLine  & "                                                                                    - BlueZone Script Team"
    testing_run = TRUE
    testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/testing_trial/notes/documents-received.vbs"
    Call run_from_GitHub(testing_script_url)
End if

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
Call changelog_update("01/03/2020", "Added new functionality to ask about accepting documents in ECF as a reminder at the end of the script.", "Casey Love, Hennepin County")
Call changelog_update("09/25/2019", "Bug Fix - script would error/stop if case was stuck in background. Added a number of checks to be sure case is not in background so the script run can continue.", "Casey Love, Hennepin County")
Call changelog_update("07/29/2019", "Bug fix - script was not identifying document information as complete when only SHEL editbox was filled.", "Casey Love, Hennepin County")
Call changelog_update("07/27/2019", "Functionality for specific forms:  Assets, MOF, AREP, LTC 1503, and MTAF. Form functionality can be accessed by checkboxes on the main dialog though all document detail can still be added in theeditboxes on the main dialog.", "Casey Love, Hennepin County")
call changelog_update("03/08/2019", "EVF received functionality added. This used to be a seperate script and will now be a part of documents received.", "Casey Love, Hennepin County")
call changelog_update("01/03/2017", "Added HSR scanner option for Hennepin County users only.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS BLOCK ===========================================================================================================

function get_footer_month_from_date(footer_month_variable, footer_year_variable, date_variable)

    footer_month_variable = DatePart("m", date_variable)
    footer_month_variable = Right("00" & footer_month_variable, 2)

    footer_year_variable = DatePart("yyyy", date_variable)
    footer_year_variable = Right(footer_year_variable, 2)

end function

function cancel_continue_confirmation(skip_functionality)

    skip_functionality = FALSE
    If ButtonPressed = 0 then       'this is the cancel button
        cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no information provided so far will be updated or noted. If you choose 'No' the update for THIS FORM will be cancelled and rest of the script will continue." & vbNewLine & vbNewLine & "YES - Stop the script entirely." & vbNewLine & "NO - Do not stop the script entrirely, just cancel the entry of this form information."& vbNewLine & "CANCEL - I didn't mean to cancel at all. (Cancel my cancel)", vbQuestion + vbYesNoCancel, "Clarify Cancel")
        If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")     'ends the script entirely
        If cancel_clarify = vbNo Then skip_functionality = TRUE
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
    End if

end function

function update_ACCT_panel_from_dialog()

    EMWriteScreen "                    ", 7, 44
    EMWriteScreen "                    ", 8, 44
    EMWriteScreen "        ", 10, 46
    EMWriteScreen "  ", 11, 44
    EMWriteScreen "  ", 11, 47
    EMWriteScreen "  ", 11, 50
    EMWriteScreen "        ", 12, 46

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 2), 6, 44
    EMWriteScreen ASSETS_ARRAY(ast_number, asset_counter), 7, 44
    EMWriteScreen ASSETS_ARRAY(ast_location, asset_counter), 8, 44
    EMWriteScreen ASSETS_ARRAY(ast_balance, asset_counter), 10, 46
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 10, 64
    Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_bal_date, asset_counter), 0, 11, 44)
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_YN, asset_counter), 12, 64
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_verif, asset_counter), 12, 72
    EMWriteScreen ASSETS_ARRAY(ast_wdrw_penlty, asset_counter), 12, 46
    EMWriteScreen ASSETS_ARRAY(apply_to_CASH, asset_counter), 14, 50
    EMWriteScreen ASSETS_ARRAY(apply_to_SNAP, asset_counter), 14, 57
    EMWriteScreen ASSETS_ARRAY(apply_to_HC, asset_counter), 14, 64
    EMWriteScreen ASSETS_ARRAY(apply_to_GRH, asset_counter), 14, 72
    EMWriteScreen ASSETS_ARRAY(apply_to_IVE, asset_counter), 14, 80
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 15, 44
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 15, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 15, 80
    If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) <> "" Then
        EMWriteScreen left(ASSETS_ARRAY(ast_next_inrst_date, asset_counter), 2), 17, 57
        EMWriteScreen right(ASSETS_ARRAY(ast_next_inrst_date, asset_counter), 2), 17, 60
    Else
        EMWriteScreen "  ", 17, 57
        EMWriteScreen "  ", 17, 60
    End If
end function

function update_SECU_panel_from_dialog()

    EMWriteScreen "            ", 7, 50
    EMWriteScreen "                    ", 8, 50
    EMWriteScreen "        ", 10, 52
    EMWriteScreen "  ", 11, 35
    EMWriteScreen "  ", 11, 38
    EMWriteScreen "  ", 11, 41
    EMWriteScreen "        ", 12, 52
    EMWriteScreen "        ", 13, 52

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 2), 6, 50
    EMWriteScreen ASSETS_ARRAY(ast_number, asset_counter), 7, 50
    EMWriteScreen ASSETS_ARRAY(ast_location, asset_counter), 8, 50
    EMWriteScreen ASSETS_ARRAY(ast_csv, asset_counter), 10, 52
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 11, 50
    Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_bal_date, asset_counter), 0, 11, 35)
    EMWriteScreen ASSETS_ARRAY(ast_face_value, asset_counter), 12, 52
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_YN, asset_counter), 13, 72
    EMWriteScreen ASSETS_ARRAY(ast_wthdr_verif, asset_counter), 13, 80
    EMWriteScreen ASSETS_ARRAY(ast_wdrw_penlty, asset_counter), 13, 52
    EMWriteScreen ASSETS_ARRAY(apply_to_CASH, asset_counter), 15, 50
    EMWriteScreen ASSETS_ARRAY(apply_to_SNAP, asset_counter), 15, 57
    EMWriteScreen ASSETS_ARRAY(apply_to_HC, asset_counter), 15, 64
    EMWriteScreen ASSETS_ARRAY(apply_to_GRH, asset_counter), 15, 72
    EMWriteScreen ASSETS_ARRAY(apply_to_IVE, asset_counter), 15, 80
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 16, 44
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 80

end function

function update_CARS_panel_from_dialog()

    EMWriteScreen "                ", 8, 43
    EMWriteScreen "                ", 8, 66
    EMWriteScreen "         ", 9, 45
    EMWriteScreen "         ", 9, 62
    EMWriteScreen "         ", 12, 45
    EMWriteScreen "  ", 13, 43
    EMWriteScreen "  ", 13, 46
    EMWriteScreen "  ", 13, 49

    EMWriteScreen left(ASSETS_ARRAY(ast_type, asset_counter), 1), 6, 43
    EMWriteScreen ASSETS_ARRAY(ast_year, asset_counter), 8, 31
    EMWriteScreen ASSETS_ARRAY(ast_make, asset_counter), 8, 43
    EMWriteScreen ASSETS_ARRAY(ast_model, asset_counter), 8, 66
    EMWriteScreen ASSETS_ARRAY(ast_trd_in, asset_counter), 9, 45
    EMWriteScreen ASSETS_ARRAY(ast_loan_value, asset_counter), 9, 62
    EMWriteScreen left(ASSETS_ARRAY(ast_value_srce, asset_counter), 1), 9, 80
    EMWriteScreen left(ASSETS_ARRAY(ast_verif, asset_counter), 1), 10, 60
    EMWriteScreen ASSETS_ARRAY(ast_amt_owed, asset_counter), 12, 45
    EMWriteScreen left(ASSETS_ARRAY(ast_owe_verif, asset_counter), 1), 12, 60
    If ASSETS_ARRAY(ast_owed_date, asset_counter) <> "" Then Call create_MAXIS_friendly_date(ASSETS_ARRAY(ast_owed_date, asset_counter), 0, 13, 43)
    EMWriteScreen left(ASSETS_ARRAY(ast_use, asset_counter), 1), 15, 43
    EMWriteScreen ASSETS_ARRAY(ast_hc_benefit, asset_counter), 15, 76
    EMWriteScreen ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter), 16, 43
    EMWriteScreen left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 76
    EMWriteScreen right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1), 16, 80

end function

'===========================================================================================================================

'DECLARATIONS ==============================================================================================================

Dim ASSETS_ARRAY()
ReDim ASSETS_ARRAY(update_panel, 0)

Const ast_panel         = 0
Const ast_owner         = 1
Const ast_ref_nbr       = 2
Const ast_instance      = 3
Const ast_type          = 4
Const ast_balance       = 5
Const ast_verif         = 6
Const ast_number        = 7
Const ast_wthdr_YN      = 8
Const ast_wdrw_penlty   = 9
Const ast_wthdr_verif   = 10
Const ast_jnt_owner_YN  = 11
Const ast_own_ratio      = 12
Const ast_othr_ownr_one = 13
Const ast_othr_ownr_two = 14
Const ast_othr_ownr_thr = 15
Const ast_owner_signed  = 16
Const apply_to_CASH     = 17
Const apply_to_SNAP     = 18
Const apply_to_HC       = 19
Const apply_to_GRH      = 20
Const apply_to_IVE      = 21
Const ast_location      = 22
Const ast_model         = 23
Const ast_make          = 24
Const ast_year          = 25
Const ast_trd_in        = 26
Const ast_loan_value    = 27
Const ast_value_srce    = 28
Const ast_amt_owed      = 29
Const ast_owe_verif     = 30
Const ast_owed_date     = 31
Const ast_hc_benefit    = 32
Const ast_bal_date      = 33
Const ast_verif_date    = 34
Const ast_next_inrst_date = 35
Const ast_owe_YN        = 36
Const ast_use           = 37
Const update_date       = 38
Const cnote_panel       = 39
Const ast_csv           = 40
Const ast_face_value    = 41
Const ast_share_note    = 42
Const ast_note          = 43

Const update_panel      = 44

Dim client_list_array

script_that_handles_documents = TRUE

'===========================================================================================================================
'Specific Forms Handled For

'EVF HANDLING
'AREP FORM HANDLING - to do
'CHANGE REPORT FORM HANDLING - to do
'LTC 1503 HANDLING - to do (this will likely call the other script)
'LTC 5181 HANDLING - to do (this will likely call the other script)
'MOF HANDLING - to do
'MSQ HANDLING - to do
'???? OHP RECEIVED HANDLING - to do - WHAT IS THIS
'ASSET FORM HANDLING - to do (no existing script for this)
'IAAs HANDLING - to do (no existing script for this)
'SHELTER FORM HANDLING - to do

'THE SCRIPT--------------------------------------------------------------------------------------------------
'dialogs on this script are embeded because there are going to be MANY dialogs


'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Call get_county_code()
end_msg = ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 136, 95, "Case number dialog"
  EditBox 65, 10, 65, 15, MAXIS_case_number
  EditBox 65, 30, 30, 15, MAXIS_footer_month
  EditBox 100, 30, 30, 15, MAXIS_footer_year
  EditBox 85, 50, 45, 15, doc_date_stamp
  ButtonGroup ButtonPressed
	OkButton 25, 75, 50, 15
	CancelButton 80, 75, 50, 15
  Text 10, 15, 50, 10, "Case number: "
  Text 10, 35, 50, 10, "Footer month:"
  Text 10, 55, 75, 10, "Document date stamp:"
  'CheckBox 10, 75, 60, 10, "OTS scanning", HSR_scanner_checkbox        'This is commented out BUT if replacing this, the dialog needs to be resized and buttons moved.'
EndDialog
Do
	DO
	    err_msg = ""
		dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
		IF IsNumeric(MAXIS_footer_month) = FALSE OR IsNumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbNewLine &  "* You must type a valid footer month and year."
        If IsDate(doc_date_stamp) = FALSE Then err_msg = err_msg & vbNewLine & "* Please enter a valid document date."
        If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

need_final_note = TRUE

'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel or VbDefaultButton2) 'defaults to no since that is most commonly chosen option
If LTC_case = vbCancel then stopscript

'Displays the dialog and navigates to case note
'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
DO
	Do
        err_msg = ""
		If LTC_case = vbNo then
			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
            BeginDialog Dialog1, 0, 0, 416, 375, "Documents received"       'This is the regular (NON LTC) dialog
              EditBox 80, 5, 330, 15, docs_rec
              EditBox 35, 50, 315, 15, ADDR
              EditBox 75, 70, 275, 15, SCHL
              EditBox 35, 90, 315, 15, DISA
              CheckBox 370, 95, 30, 10, "MOF", mof_form_checkbox
              EditBox 35, 110, 315, 15, JOBS
              CheckBox 370, 115, 30, 10, "EVF", evf_form_received_checkbox
              EditBox 35, 130, 315, 15, BUSI
              EditBox 35, 150, 315, 15, UNEA
              EditBox 35, 170, 315, 15, ACCT
              CheckBox 370, 175, 30, 10, "Asset", asset_form_checkbox
              Text 370, 185, 35, 10, "Statement"
              EditBox 60, 190, 290, 15, other_assets
              CheckBox 370, 200, 30, 10, "AREP", arep_form_checkbox
              EditBox 35, 210, 315, 15, SHEL
              EditBox 35, 230, 315, 15, INSA
              EditBox 55, 250, 295, 15, other_verifs
              CheckBox 370, 255, 30, 10, "MTAF", mtaf_form_checkbox
              EditBox 80, 290, 320, 15, notes
              EditBox 80, 310, 320, 15, actions_taken
              EditBox 80, 330, 320, 15, verifs_needed
              EditBox 220, 355, 80, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 305, 355, 50, 15
                CancelButton 360, 355, 50, 15
              Text 10, 95, 25, 10, "DISA:"
              Text 10, 115, 25, 10, "JOBS:"
              Text 10, 135, 20, 10, "BUSI:"
              Text 10, 155, 25, 10, "UNEA:"
              Text 10, 175, 25, 10, "ACCT:"
              Text 10, 195, 45, 10, "Other assets:"
              Text 10, 215, 25, 10, "SHEL:"
              Text 10, 235, 20, 10, "INSA:"
              Text 10, 255, 45, 10, "Other verif's:"
              Text 155, 360, 60, 10, "Worker signature:"
              Text 10, 55, 25, 10, "ADDR:"
              Text 10, 295, 70, 10, "Notes on your doc's:"
              Text 10, 315, 50, 10, "Actions taken:"
              Text 140, 25, 205, 10, "Note: What you enter above will become the case note header."
              Text 10, 10, 70, 10, "Documents received: "
              Text 10, 335, 65, 10, "Verif's still needed:"
              GroupBox 5, 35, 350, 235, "Breakdown of Documents received"
              GroupBox 5, 275, 405, 75, "Additional information"
              Text 10, 75, 65, 10, "SCHL/STIN/STEC:"
              GroupBox 360, 35, 50, 235, "FORMS"
              Text 370, 45, 35, 45, "Watch for more form options - coming soon!"
            EndDialog

        ElseIf LTC_case = vbYes then
			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
            BeginDialog Dialog1, 0, 0, 416, 405, "Documents received LTC"           'This is the LTC Dialog
              EditBox 80, 5, 330, 15, docs_rec
              EditBox 35, 45, 315, 15, FACI
              EditBox 35, 65, 135, 15, JOBS
              EditBox 215, 65, 135, 15, BUSI_RBIC
              CheckBox 370, 70, 30, 10, "EVF", evf_form_received_checkbox
              EditBox 35, 85, 315, 15, UNEA
              EditBox 35, 105, 315, 15, ACCT
              CheckBox 370, 110, 30, 10, "Asset", asset_form_checkbox
              Text 370, 120, 35, 10, "Documents"
              EditBox 35, 125, 315, 15, SECU
              EditBox 35, 145, 315, 15, CARS
              EditBox 35, 165, 315, 15, REST
              EditBox 65, 185, 285, 15, OTHR
              EditBox 35, 205, 315, 15, SHEL
              EditBox 35, 225, 315, 15, INSA
              EditBox 80, 245, 270, 15, medical_expenses
              CheckBox 370, 260, 30, 10, "AREP", arep_form_checkbox
              EditBox 55, 265, 295, 15, veterans_info
              CheckBox 365, 275, 40, 10, "LTC1503", ltc_1503_form_checkbox
              EditBox 55, 285, 295, 15, other_verifs
              EditBox 80, 320, 330, 15, notes
              EditBox 80, 340, 330, 15, actions_taken
              EditBox 80, 360, 330, 15, verifs_needed
              EditBox 225, 385, 80, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 310, 385, 50, 15
                CancelButton 360, 385, 50, 15
              Text 10, 150, 20, 10, "CARS:"
              Text 10, 170, 20, 10, "REST:"
              Text 10, 190, 50, 10, "BURIAL/OTHR:"
              Text 10, 210, 25, 10, "SHEL:"
              Text 10, 230, 25, 10, "INSA:"
              Text 10, 290, 45, 10, "Other verif's:"
              Text 165, 390, 60, 10, "Worker signature:"
              Text 10, 50, 25, 10, "FACI:"
              Text 10, 325, 70, 10, "Notes on your doc's:"
              Text 10, 345, 50, 10, "Actions taken:"
              Text 205, 20, 205, 10, "Note: What you enter above will become the case note header."
              Text 5, 10, 70, 10, "Documents received: "
              Text 10, 365, 70, 10, "Verif's still needed:"
              GroupBox 5, 30, 350, 275, "Breakdown of Documents received"
              Text 10, 110, 20, 10, "ACCT:"
              Text 175, 70, 40, 10, "BUSI/RBIC:"
              Text 10, 90, 25, 10, "UNEA:"
              Text 10, 270, 45, 10, "Veteran info:"
              Text 10, 70, 20, 10, "JOBS:"
              Text 10, 250, 65, 10, "Medical expenses:"
              GroupBox 5, 310, 410, 70, "Additional information"
              Text 10, 130, 20, 10, "SECU:"
              GroupBox 360, 35, 50, 275, "FORMS"
              Text 370, 145, 35, 45, "Watch for more form options - coming soon!"
            EndDialog
        End If

        dialog Dialog1
		cancel_confirmation																'quits if cancel is pressed

		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* You must sign your case note."

        no_checkboxes_checked = FALSE
        no_detail_added = False
        If LTC_case = vbNo then
            If mof_form_checkbox = unchecked AND evf_form_received_checkbox = unchecked AND asset_form_checkbox = unchecked AND arep_form_checkbox = uncehecked AND mtaf_form_checkbox = unchecked Then no_checkboxes_checked = TRUE
            If trim(ADDR) = "" AND trim(SCHL) = "" AND trim(DISA) = "" AND trim(JOBS) = "" AND trim(BUSI) = "" AND trim(UNEA) = "" AND trim(ACCT) = "" AND trim(other_assets) = "" AND trim(SHEL) = "" AND trim(INSA) = "" AND trim(other_verifs) = "" Then no_detail_added = TRUE
        ElseIf LTC_case = vbYes then
            If evf_form_received_checkbox = unchecked AND asset_form_checkbox = unchecked AND arep_form_checkbox = uncehecked AND ltc_1503_form_checkbox = unchecked Then no_checkboxes_checked = TRUE
            If trim(FACI) = "" AND trim(JOBS) = "" AND trim(BUSI_RBIC) = "" AND trim(UNEA) = "" AND trim(ACCT) = "" AND trim(SECU) = "" AND trim(CARS) = "" AND trim(REST) = "" AND trim(OTHR) = "" AND trim(SHEL) = "" AND trim(INSA) = "" AND trim(medical_expenses) = "" AND trim(veterans_info) = "" AND trim(other_verifs) = "" Then no_detail_added = TRUE
        End If

        If no_checkboxes_checked = TRUE AND no_detail_added = TRUE Then err_msg = err_msg & vbNewLine & "* Detail about the documents relation to the case and potential changes/updates should be listed by the most appropriate panel type, or select one of the forms if one was received to have specific functionality run for that form."
        ' If HSR_scanner_checkbox = unchecked and actions_taken = "" Then
        '     If ltc_1503_form_checkbox = unchecked AND evf_form_received_checkbox = unchecked AND asset_form_checkbox = unchecked AND mof_form_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* You must case note your actions taken."
        ' End If

        If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

	LOOP until err_msg = ""													'Loops until that case number exists
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

Call MAXIS_background_check

CALL Generate_Client_List(client_dropdown, "Select One...")
CALL Generate_Client_List(client_dropdown_CB, "Select or Type")

If LTC_case = vbNo then end_msg = "Success! Documents received noted for case."
If LTC_case = vbYes then end_msg = "Success! Documents received noted for LTC case."

'This will be for any functionality that needs the HH Member array
If asset_form_checkbox = checked Then
    call HH_member_custom_dialog(HH_member_array)
End If

'EVF HANDLING =======================================================================================
If evf_form_received_checkbox = checked Then
    EVF_TIKL_checkbox = checked 'defaulting the TIKL checkbox to be checked initially in the dialog.
    evf_date_recvd = doc_date_stamp
    'starts the EVF received case note dialog
    DO
    	Do
		    '-------------------------------------------------------------------------------------------------DIALOG
		    Dialog1 = "" 'Blanking out previous dialog detail
		    BeginDialog Dialog1, 0, 0, 291, 205, "Employment Verification Form Received"
		      Text 70, 10, 60, 10, MAXIS_case_number
		      EditBox 220, 5, 60, 15, evf_date_recvd
		      ComboBox 70, 30, 210, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
		      EditBox 70, 50, 210, 15, employer
		      DropListBox 70, 75, 210, 45, client_dropdown, evf_client
		      DropListBox 75, 110, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", info
		      EditBox 220, 110, 60, 15, info_date
		      EditBox 75, 130, 60, 15, request_info
		      CheckBox 160, 135, 105, 10, "Create TIKL for additional info", EVF_TIKL_checkbox
		      EditBox 70, 160, 210, 15, actions_taken
		      ButtonGroup ButtonPressed
		    	OkButton 175, 180, 50, 15
		    	CancelButton 230, 180, 50, 15
		      Text 10, 135, 65, 10, "Info Requested via:"
		      Text 10, 115, 60, 10, "Addt'l Info Reqstd:"
		      Text 5, 75, 60, 10, "Household Memb:"
		      Text 10, 55, 55, 10, "Employer name:"
		      Text 15, 165, 50, 10, "Actions taken:"
		      Text 25, 35, 40, 10, "EVF Status:"
		      Text 150, 10, 65, 10, "Date EVF received:"
		      Text 15, 10, 50, 10, "Case Number:"
		      Text 160, 115, 55, 10, "Date Requested:"
		      GroupBox 5, 95, 280, 60, "Is additional information needed?"
		    EndDialog
    		err_msg = ""
    		Dialog Dialog1       	'starts the EVF dialog
            Call cancel_continue_confirmation(skip_evf)
    		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You need to type a valid case number."
    		IF IsDate(evf_date_recvd) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date for date the EVF was received."
    		If EVF_status_dropdown = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select the status of the EVF on the dropdown menu"		'checks that there is a date in the date received box
    		IF employer = "" THEN err_msg = err_msg & vbCr & "* You must enter the employers name."  'checks if the employer name has been entered
    		IF evf_client = "Select One..." THEN err_msg = err_msg & vbCr & "* You must enter the MEMB information."  'checks if the client name has been entered
    		IF info = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select if additional info was requested."  'checks if completed by employer was selected
    		IF info = "yes" and IsDate(info_date) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date that additional info was requested."  'checks that there is a info request date entered if the it was requested
    		IF info = "yes" and request_info = "" THEN err_msg = err_msg & vbCr & "* You must enter the method used to request additional info."		'checks that there is a method of inquiry entered if additional info was requested
    		If info = "no" and request_info <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have information requested."
    		If info = "no" and info_date <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have a date requested."
    		If EVF_TIKL_checkbox = 1 and info <> "yes" then err_msg = err_msg & vbCr & "* Additional informaiton was not requested, uncheck the TIKL checkbox."
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_evf = TRUE Then
                evf_form_received_checkbox = unchecked
                err_msg = ""
                EVF_TIKL_checkbox = unchecked
            End If

    		IF err_msg <> "" AND left(err_msg,4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "* Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
End If

if evf_form_received_checkbox = checked Then
    evf_ref_numb = left(evf_client, 2)
    docs_rec = docs_rec & ", EVF for M" & evf_ref_numb

    'Checks if additional info is yes and the TIKL is checked, sets a TIKL for the return of the info
    If EVF_TIKL_checkbox = checked Then
        'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        Call create_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action.", 10, date, True, TIKL_note_text)
    	'Success message
    	end_msg = end_msg & vbNewLine & "Additional detail added about EVF." & vbNewLine & "TIKL has been sent for 10 days from now for the additional information requested."
    Else
        end_msg = end_msg & vbNewLine & "Additional detail added about EVF."
    End If
End If

If mof_form_checkbox = checked Then
    mof_date_recd = doc_date_stamp
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 226, 255, "Medical Opinion Form Received for Case # " & MAXIS_case_number
      EditBox 55, 5, 50, 15, mof_date_recd
      CheckBox 125, 10, 85, 10, "Client signed release?", mof_clt_release_checkbox
      DropListBox 80, 25, 140, 45, client_dropdown, mof_hh_memb
      EditBox 90, 45, 55, 15, last_exam_date
      EditBox 90, 65, 55, 15, doctor_date
      ComboBox 70, 85, 150, 45, "Select or Type"+chr(9)+"Less than 30 Days"+chr(9)+"Between 30 - 45 Days"+chr(9)+"More than 45 Days"+chr(9)+"No End Date Listed", mof_time_condition_will_last
      EditBox 85, 105, 135, 15, ability_to_work
      EditBox 55, 125, 165, 15, mof_other_notes
      EditBox 55, 145, 165, 15, actions_taken
      CheckBox 10, 170, 215, 10, "Check here if the MOF indicates an SSA application is needed.", SSA_application_indicated_checkbox
      CheckBox 10, 185, 185, 10, "Check here if DISA will be updated as needed by TTL", TTL_to_update_checkbox
      CheckBox 10, 200, 190, 10, "Check here if you sent an email to TTL/FSS DataTeam.", TTL_email_checkbox
      EditBox 90, 215, 65, 15, TTL_email_date
      ButtonGroup ButtonPressed
        OkButton 115, 235, 50, 15
        CancelButton 170, 235, 50, 15
      Text 5, 10, 50, 10, "Date received: "
      Text 5, 30, 70, 10, "HHLD Member name"
      Text 5, 50, 60, 10, "Date of last exam: "
      Text 5, 70, 80, 10, "Date doctor signed form: "
      Text 155, 45, 50, 35, "Do not enter diagnosis in case notes per PQ #16506."
      Text 5, 90, 60, 10, "Condition will last:"
      Text 5, 110, 75, 10, "Client's ability to work: "
      Text 5, 130, 40, 10, "Other notes: "
      Text 5, 150, 45, 10, "Action taken: "
      Text 30, 220, 55, 10, "Date email sent:"
    EndDialog

    Do
        DO
            Err_msg = ""
            Dialog Dialog1
            Call cancel_continue_confirmation(skip_MOF)
            'Call validate_MAXIS_case_number(err_msg, "*")
            If IsDate(mof_date_recd) = FALSE Then err_msg = err_ms & vbNewLine & "* Enter a valid date the document was received."
            If mof_hh_memb = "Select One..." Then err_msg = err_ms & vbNewLine & "* Select the household member."
            IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered
            If TTL_email_checkbox = checked Then
                If IsDate(TTL_email_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the date an email about this MOF was sent to TTL."
            End If
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_MOF= TRUE Then
                err_msg = ""
                mof_form_checkbox = unchecked
            End If
            If err_msg <> "" Then msgbox "Please resolve the following for the script to continue:" & vbNewLine & err_msg
        LOOP until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    last_exam_date = trim(last_exam_date)
    doctor_date = trim(doctor_date)
    If mof_time_condition_will_last = "Select or Type" Then mof_time_condition_will_last = ""
    mof_time_condition_will_last = trim(mof_time_condition_will_last)
    ability_to_work = trim(ability_to_work)
    mof_other_notes = trim(mof_other_notes)
End If

If mof_form_checkbox = checked Then
    mof_ref_numb = left(mof_hh_memb, 2)
    docs_rec = docs_rec & ", MOF for M" & mof_ref_numb
    end_msg = end_msg & vbNewLine & "Additional detail about MOF."
End If


If asset_form_checkbox = checked Then
    asset_counter = 0
    skip_asset = FALSE
    Call navigate_to_MAXIS_screen("STAT", "ACCT")
    For each member in HH_member_array
        Call write_value_and_transmit(member, 20, 76)

        EMReadScreen acct_versions, 1, 2, 78
        If acct_versions <> "0" Then
            EMWriteScreen "01", 20, 79
            transmit
            Do
                EMReadScreen ACCT_instance, 1, 2, 73
                EMReadScreen ACCT_type, 2, 6, 44
                EMReadScreen ACCT_nbr, 20, 7, 44
                EMReadScreen ACCT_location, 20, 8, 44
                EMReadScreen ACCT_balance, 8, 10, 46
                EMReadScreen ACCT_bal_verif, 1, 10, 64
                EMReadScreen ACCT_bal_date, 8, 11, 44
                EMReadScreen ACCT_withdraw_pen, 8, 12, 46
                EMReadScreen ACCT_withdraw_YN, 1, 12, 64
                EMReadScreen ACCT_withdraw_verif, 1, 12, 72
                EMReadScreen ACCT_cash, 1, 14, 50
                EMReadScreen ACCT_snap, 1, 14, 57
                EMReadScreen ACCT_hc, 1, 14, 64
                EMReadScreen ACCT_grh, 1, 14, 72
                EMReadScreen ACCT_ive, 1, 14, 80
                EMReadScreen ACCT_joint_owner_YN, 1, 15, 44
                EMReadScreen ACCT_share_ratio, 5, 15, 76
                EMReadScreen ACCT_next_interest, 5, 17, 57
                EMReadScreen ACCT_updated_date, 8, 21, 55

                ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

                ASSETS_ARRAY(ast_panel, asset_counter) = "ACCT"
                ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
                For each person in client_list_array
                    If left(person, 2) = member then
                        ASSETS_ARRAY(ast_owner, asset_counter) = person
                        Exit For
                    End If
                Next
                ASSETS_ARRAY(ast_instance, asset_counter) = "0" & ACCT_instance
                If ACCT_type = "SV" Then ASSETS_ARRAY(ast_type, asset_counter) = "SV - Savings"
                If ACCT_type = "CK" Then ASSETS_ARRAY(ast_type, asset_counter) = "CK - Checking"
                If ACCT_type = "CD" Then ASSETS_ARRAY(ast_type, asset_counter) = "CD - Cert of Deposit"
                If ACCT_type = "MM" Then ASSETS_ARRAY(ast_type, asset_counter) = "MM - Money market"
                If ACCT_type = "DC" Then ASSETS_ARRAY(ast_type, asset_counter) = "DC - Debit Card"
                If ACCT_type = "KO" Then ASSETS_ARRAY(ast_type, asset_counter) = "KO - Keogh Account"
                If ACCT_type = "FT" Then ASSETS_ARRAY(ast_type, asset_counter) = "FT - Federatl Thrift SV plan"
                If ACCT_type = "SL" Then ASSETS_ARRAY(ast_type, asset_counter) = "SL - Stat/Local Govt Ret"
                If ACCT_type = "RA" Then ASSETS_ARRAY(ast_type, asset_counter) = "RA - Employee Ret Annuities"
                If ACCT_type = "NP" Then ASSETS_ARRAY(ast_type, asset_counter) = "NP - Non-Profit Employer Ret Plan"
                If ACCT_type = "IR" Then ASSETS_ARRAY(ast_type, asset_counter) = "IR - Indiv Ret Acct"
                If ACCT_type = "RH" Then ASSETS_ARRAY(ast_type, asset_counter) = "RH - Roth IRA"
                If ACCT_type = "FR" Then ASSETS_ARRAY(ast_type, asset_counter) = "FR - Ret Plans for Employers"
                If ACCT_type = "CT" Then ASSETS_ARRAY(ast_type, asset_counter) = "CT - Corp Ret Trust"
                If ACCT_type = "RT" Then ASSETS_ARRAY(ast_type, asset_counter) = "RT - Other Ret Fund"
                If ACCT_type = "QT" Then ASSETS_ARRAY(ast_type, asset_counter) = "QT - Qualified Tuition (529)"
                If ACCT_type = "CA" Then ASSETS_ARRAY(ast_type, asset_counter) = "CA - Coverdell SV (530)"
                If ACCT_type = "OE" Then ASSETS_ARRAY(ast_type, asset_counter) = "OE - Other Educational "
                If ACCT_type = "OT" Then ASSETS_ARRAY(ast_type, asset_counter) = "OT - Other"
                ASSETS_ARRAY(ast_number, asset_counter) = replace(ACCT_nbr, "_", "")
                ASSETS_ARRAY(ast_location, asset_counter) = replace(ACCT_location, "_", "")
                ASSETS_ARRAY(ast_balance, asset_counter) = trim(ACCT_balance)
                If ACCT_bal_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1 - Bank Statement"
                If ACCT_bal_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - Agcy Ver Form"
                If ACCT_bal_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - Coltrl Document"
                If ACCT_bal_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
                If ACCT_bal_verif = "6" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"
                If ACCT_bal_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prvd"
                ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(ACCT_bal_date, " ", "/")
                If ASSETS_ARRAY(ast_bal_date, asset_counter) = "__/__/__" Then ASSETS_ARRAY(ast_bal_date, asset_counter) = ""
                ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(replace(ACCT_withdraw_pen, "_", ""))
                ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = replace(ACCT_withdraw_YN, "_", "")
                ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = replace(ACCT_withdraw_verif, "_", "")
                ASSETS_ARRAY(apply_to_CASH, asset_counter) = replace(ACCT_cash, "_", "")
                ASSETS_ARRAY(apply_to_SNAP, asset_counter) = replace(ACCT_snap, "_", "")
                ASSETS_ARRAY(apply_to_HC, asset_counter) = replace(ACCT_hc, "_", "")
                ASSETS_ARRAY(apply_to_GRH, asset_counter) = replace(ACCT_grh, "_", "")
                ASSETS_ARRAY(apply_to_IVE, asset_counter) = replace(ACCT_ive, "_", "")
                ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = replace(ACCT_joint_owner_YN, "_", "")
                ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(ACCT_share_ratio, " ", "")
                ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = replace(ACCT_next_interest, " ", "/")
                If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = "__/__" Then ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = ""
                ASSETS_ARRAY(update_panel, asset_counter) = unchecked
                ASSETS_ARRAY(update_date, asset_counter) = replace(ACCT_updated_date, " ", "/")

                transmit
                asset_counter = asset_counter + 1
                'MsgBox asset_counter
                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
            Loop until reached_last_ACCT_panel = "ENTER A VALID"
        End If
    Next

    Call navigate_to_MAXIS_screen("STAT", "SECU")
    For each member in HH_member_array
        Call write_value_and_transmit(member, 20, 76)

        EMReadScreen secu_versions, 1, 2, 78
        If secu_versions <> "0" Then
            EMWriteScreen "01", 20, 79
            transmit
            Do

                EMReadScreen SECU_instance, 1, 2, 73
                EMReadScreen SECU_type, 2, 6, 50
                EMReadScreen SECU_acct_number, 12, 7, 50
                EMReadScreen SECU_name, 20, 8, 50
                EMReadScreen SECU_csv, 8, 10, 52
                EMReadScreen SECU_value_date, 8, 11, 35
                EMReadScreen SECU_verif, 1, 11, 50
                EMReadScreen SECU_face_value, 8, 12, 52
                EMReadScreen SECU_withdraw_amount, 8, 13, 52
                EMReadScreen SECU_wthdrw_YN, 1, 13, 72
                EMReadScreen SECU_wthdrw_verif, 1, 13, 80
                EMReadScreen SECU_apply_to_CASH, 1, 15, 50
                EMReadScreen SECU_apply_to_SNAP, 1, 15, 57
                EMReadScreen SECU_apply_to__HC, 1, 15, 64
                EMReadScreen SECU_apply_to_GRH, 1, 15, 72
                EMReadScreen SECU_apply_to_IVE, 1, 15, 80
                EMReadScreen SECU_joint_owner_YN, 1, 16, 44
                EMReadScreen SECU_share_ratio, 5, 16, 76
                EMReadScreen SECU_updated_date, 8, 21, 55


                ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

                ASSETS_ARRAY(ast_panel, asset_counter) = "SECU"
                ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
                For each person in client_list_array
                    If left(person, 2) = member then
                        ASSETS_ARRAY(ast_owner, asset_counter) = person
                        Exit For
                    End If
                Next
                ASSETS_ARRAY(ast_instance, asset_counter) = "0" & SECU_instance
                If SECU_type = "LI" Then ASSETS_ARRAY(ast_type, asset_counter) = "LI - Life Insurance"
                If SECU_type = "ST" Then ASSETS_ARRAY(ast_type, asset_counter) = "ST - Stocks"
                If SECU_type = "BO" Then ASSETS_ARRAY(ast_type, asset_counter) = "BO - Bonds"
                If SECU_type = "CD" Then ASSETS_ARRAY(ast_type, asset_counter) = "CD - Ctrct For Deed"
                If SECU_type = "MO" Then ASSETS_ARRAY(ast_type, asset_counter) = "MO - Mortgage Note"
                If SECU_type = "AN" Then ASSETS_ARRAY(ast_type, asset_counter) = "AN - Annuity"
                If SECU_type = "OT" Then ASSETS_ARRAY(ast_type, asset_counter) = "OT - Other"
                ASSETS_ARRAY(ast_number, asset_counter) = replace(SECU_acct_number, "_", "")
                ASSETS_ARRAY(ast_location, asset_counter) = replace(SECU_name, "_", "")
                ASSETS_ARRAY(ast_csv, asset_counter) = trim(SECU_csv)
                ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(SECU_value_date, " ", "/")
                If SECU_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1  - Agency Form"
                If SECU_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - Source Doc"
                If SECU_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - Phone Contact"
                If SECU_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
                If SECU_verif = "6" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Stmt"
                If SECU_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prov"
                ASSETS_ARRAY(ast_face_value, asset_counter) = replace(trim(SECU_face_value), "_", "")
                ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(replace(SECU_withdraw_amount, "_", ""))
                ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = replace(SECU_wthdrw_YN, "_", "")
                ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = replace(SECU_wthdrw_verif, "_", "")
                ASSETS_ARRAY(apply_to_CASH, asset_counter) = replace(SECU_apply_to_CASH, "_", "")
                ASSETS_ARRAY(apply_to_SNAP, asset_counter) = replace(SECU_apply_to_SNAP, "_", "")
                ASSETS_ARRAY(apply_to_HC, asset_counter) = replace(SECU_apply_to_HC, "_", "")
                ASSETS_ARRAY(apply_to_GRH, asset_counter) = replace(SECU_apply_to_GRH, "_", "")
                ASSETS_ARRAY(apply_to_IVE, asset_counter) = replace(SECU_apply_to_IVE, "_", "")
                ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = replace(SECU_joint_owner_YN, "_", "")
                ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(SECU_share_ratio, " ", "")
                ASSETS_ARRAY(update_date, asset_counter) = replace(SECU_updated_date, " ", "/")
                ASSETS_ARRAY(update_panel, asset_counter) = Unchecked

                transmit
                asset_counter = asset_counter + 1
                EMReadScreen reached_last_SECU_panel, 13, 24, 2
            Loop until reached_last_SECU_panel = "ENTER A VALID"
        End If
    Next

    Call navigate_to_MAXIS_screen("STAT", "CARS")
    For each member in HH_member_array
        Call write_value_and_transmit(member, 20, 76)

        EMReadScreen cars_versions, 1, 2, 78
        If cars_versions <> "0" Then
            EMWriteScreen "01", 20, 79
            transmit
            Do

                EMReadScreen CARS_instance, 1, 2, 73
                EMReadScreen CARS_type, 1, 6, 43
                EMReadScreen CARS_year, 4, 8, 31
                EMReadScreen CARS_make, 15, 8, 43
                EMReadScreen CARS_model, 15, 8, 66
                EMReadScreen CARS_trade_in, 8, 9, 45
                EMReadScreen CARS_loan, 8, 9, 62
                EMReadScreen CARS_source, 1, 9, 80
                EMReadScreen CARS_owner_verif, 1, 10, 60
                EMReadScreen CARS_owe_amount, 8, 12, 45
                EMReadScreen CARS_owed_verif, 1, 12, 60
                EMReadScreen CARS_owed_date, 8, 13, 43
                EMReadScreen CARS_use, 1, 15, 43
                EMReadScreen CARS_hc_benefit, 1, 15, 76
                EMReadScreen CARS_joint_owner_YN, 1, 16, 43
                EMReadScreen CARS_share_ratio, 5, 16, 76
                EMReadScreen CARS_updated_date, 8, 21, 55

                ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)

                ASSETS_ARRAY(ast_panel, asset_counter) = "CARS"
                ASSETS_ARRAY(ast_ref_nbr, asset_counter) = member
                For each person in client_list_array
                    If left(person, 2) = member then
                        ASSETS_ARRAY(ast_owner, asset_counter) = person
                        Exit For
                    End If
                Next
                ASSETS_ARRAY(ast_instance, asset_counter) = "0" & CARS_instance
                If CARS_type = "1" Then ASSETS_ARRAY(ast_type, asset_counter) = "1 - Car"
                If CARS_type = "2" Then ASSETS_ARRAY(ast_type, asset_counter) = "2 - Truck"
                If CARS_type = "3" Then ASSETS_ARRAY(ast_type, asset_counter) = "3 - Van"
                If CARS_type = "4" Then ASSETS_ARRAY(ast_type, asset_counter) = "4 - Camper"
                If CARS_type = "5" Then ASSETS_ARRAY(ast_type, asset_counter) = "5 - Motorcycle"
                If CARS_type = "6" Then ASSETS_ARRAY(ast_type, asset_counter) = "6 - Trailer"
                If CARS_type = "7" Then ASSETS_ARRAY(ast_type, asset_counter) = "7 - Other"
                ASSETS_ARRAY(ast_year, asset_counter) = CARS_year
                ASSETS_ARRAY(ast_make, asset_counter) = replace(CARS_make, "_", "")
                ASSETS_ARRAY(ast_model, asset_counter) = replace(CARS_model, "_", "")
                ASSETS_ARRAY(ast_trd_in, asset_counter) = trim(CARS_trade_in)
                ASSETS_ARRAY(ast_loan_value, asset_counter) = trim(CARS_loan)
                If CARS_source = "1" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "1 - NADA"
                If CARS_source = "2" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "2 - Appraisal Val"
                If CARS_source = "3" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "3 - Client Stmt"
                If CARS_source = "4" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "4 - Other Document"
                If CARS_owner_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "1 - Title"
                If CARS_owner_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "2 - License Reg"
                If CARS_owner_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "3 - DMV"
                If CARS_owner_verif = "4" Then ASSETS_ARRAY(ast_verif, asset_counter) = "4 - Purchase Agmt"
                If CARS_owner_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
                If CARS_owner_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "N - No Ver Prvd"
                ASSETS_ARRAY(ast_amt_owed, asset_counter) = trim(replace(CARS_owe_amount, "_", ""))
                ASSETS_ARRAY(ast_owe_YN, asset_counter) = replace(CARS_joint_owner_YN, "_", "")
                ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(CARS_owed_date, " ", "/")
                If ASSETS_ARRAY(ast_bal_date, asset_counter) = "__/__/__" Then ASSETS_ARRAY(ast_bal_date, asset_counter) = ""
                If CARS_use = "1" Then ASSETS_ARRAY(ast_use, asset_counter) = "1 -  Primary Veh"
                If CARS_use = "2" Then ASSETS_ARRAY(ast_use, asset_counter) = "2 - Emp/Trng Trans/Emp Search"
                If CARS_use = "3" Then ASSETS_ARRAY(ast_use, asset_counter) = "3 - Disa Trans"
                If CARS_use = "4" Then ASSETS_ARRAY(ast_use, asset_counter) = "4 - Inc Producing"
                If CARS_use = "5" Then ASSETS_ARRAY(ast_use, asset_counter) = "5 - Used As Home"
                If CARS_use = "7" Then ASSETS_ARRAY(ast_use, asset_counter) = "7 - Unlicensed"
                If CARS_use = "8" Then ASSETS_ARRAY(ast_use, asset_counter) = "8 - Othr Countable"
                If CARS_use = "9" Then ASSETS_ARRAY(ast_use, asset_counter) = "9 - Unavailable"
                If CARS_use = "0" Then ASSETS_ARRAY(ast_use, asset_counter) = "0 - Long Distance Emp Travel"
                If CARS_use = "A" Then ASSETS_ARRAY(ast_use, asset_counter) = "A - Carry Heating Fuel Or Water"
                ASSETS_ARRAY(ast_hc_benefit, asset_counter) = CARS_hc_benefit
                ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = CARS_joint_owner_YN
                ASSETS_ARRAY(ast_own_ratio, asset_counter) = replace(CARS_share_ratio, " ", "")
                ASSETS_ARRAY(update_date, asset_counter) = replace(CARS_updated_date, " ", "/")
                ASSETS_ARRAY(update_panel, asset_counter) = unchecked

                transmit
                asset_counter = asset_counter + 1
                EMReadScreen reached_last_CARS_panel, 13, 24, 2
            Loop until reached_last_CARS_panel = "ENTER A VALID"
        End If
    Next

    If LTC_case = vbNo then

        asset_form_doc_date = doc_date_stamp

        current_asset_panel = FALSE
        acct_panels = 0
        secu_panels = 0
        cars_panels = 0
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                current_asset_panel = TRUE
                acct_panels = acct_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
                asset_display = asset_display & vbNewLine & "ACCT - " & the_asset
            ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                current_asset_panel = TRUE
                secu_panels = secu_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
                asset_display = asset_display & vbNewLine & "SECU - " & the_asset
            ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                current_asset_panel = TRUE
                cars_panels = cars_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
                asset_display = asset_display & vbNewLine & "CARS - " & the_asset
            Else
                asset_display = asset_display & vbNewLine & ASSETS_ARRAY(ast_panel, the_asset) & " - " & the_asset
            End If
        Next

        'MsgBox asset_display

        dlg_len = 260

        If acct_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * acct_panels)
        If secu_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * secu_panels)
        If cars_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * cars_panels)
        'MsgBox dlg_len

		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
        y_pos = 60
        BeginDialog Dialog1, 0, 0, 390, dlg_len, "Signed Personal Statement about Assest for Case #" & MAXIS_case_number
          Text 10, 10, 265, 10, "Assets for SNAP/Cash are self attested and are reported on this form (DHS 6054)."
          Text 10, 30, 95, 10, "Date the form was received:"
          EditBox 110, 25, 35, 15, asset_form_doc_date
          Text 10, 45, 50, 10, "Action Taken:"
          EditBox 60, 40, 325, 15, actions_taken
          If acct_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current ACCT panel details."
              Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                      Text 15, y_pos, 275, 10,  "* ACCT " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset) & " - Balance: $" & ASSETS_ARRAY(ast_balance, the_asset)
                      CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
                      y_pos = y_pos + 10
                  End If
              Next
              y_pos = y_pos + 5
          End If

          Text 10, y_pos, 280, 10, "Information provided about Bank Accounts, Debit Accounts, or Certificates of Deposit:"
          y_pos = y_pos + 15
          EditBox 15, y_pos, 370, 15, box_one_info
          y_pos = y_pos + 20

          If secu_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current SECU panel details."
              Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                      Text 15, y_pos, 275, 10, "* SECU " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset)
                      CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
                      y_pos = y_pos + 10
                  End If
              Next
              y_pos = y_pos + 5
          End If
          Text 10, y_pos, 250, 10, "Information provided aboutStocks, Bonds, Pensions, or Retirement Accounts:"
          y_pos = y_pos + 15
          EditBox 15, y_pos, 370, 15, box_two_info
          y_pos = y_pos + 20
          If cars_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current CARS panel details."
              Text 260, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                      Text 15, y_pos, 275, 10, "* CARS " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset)
                      CheckBox 300, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
                      y_pos = y_pos + 10
                  End If
              Next
              y_pos = y_pos + 5
          End If
          Text 10, y_pos, 125, 10, "Information provided about Vehicles:"
          y_pos = y_pos + 15
          EditBox 15, y_pos, 370, 15, box_three_info
          y_pos = y_pos + 25
          y_pos_over = y_pos
          Text 10, y_pos, 40, 10, "Signed by:"
          Text 135, y_pos, 35, 10, "On (date):"
          y_pos = y_pos + 15
          ComboBox 10, y_pos, 105, 45, client_dropdown_CB, signed_by_one
          EditBox 130, y_pos, 50, 15, signed_one_date
          y_pos = y_pos + 20
          ComboBox 10, y_pos, 105, 45, client_dropdown_CB, signed_by_two
          EditBox 130, y_pos, 50, 15, signed_two_date
          y_pos = y_pos + 20
          ComboBox 10, y_pos, 105, 45, client_dropdown_CB, signed_by_three
          EditBox 130, y_pos, 50, 15, signed_three_date

          CheckBox 240, y_pos_over, 130, 10, "Check here to have the script update", run_updater_checkbox
          Text 250, y_pos_over + 10, 115, 10, "asset panels. (ACCT, SECU, CARS)."
          Text 255, y_pos_over + 25, 85, 20, "*Panels updated by the script will be case noted."
          ButtonGroup ButtonPressed
            OkButton 280, y_pos, 50, 15
            CancelButton 335, y_pos, 50, 15
        EndDialog

        Do
            Do
                err_msg = ""
                dialog Dialog1
                Call cancel_continue_confirmation(skip_asset)
                IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered
                If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
                If skip_asset= TRUE Then
                    err_msg = ""
                    asset_form_checkbox = unchecked
                End If
                If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop Until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE
    Else
        asset_form_doc_date = doc_date_stamp
        current_asset_panel = FALSE
        acct_panels = 0
        secu_panels = 0
        cars_panels = 0
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                current_asset_panel = TRUE
                acct_panels = acct_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then
                    ASSETS_ARRAY(ast_verif_date, the_asset) = doc_date_stamp
                    actions_taken =  actions_taken & "Updated ACCT " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & ", "
                End If
                asset_display = asset_display & vbNewLine & "ACCT - " & the_asset
            ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                current_asset_panel = TRUE
                secu_panels = secu_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then
                    ASSETS_ARRAY(ast_verif_date, the_asset) = doc_date_stamp
                    actions_taken =  actions_taken & "Updated SECU " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & ", "
                End If
                asset_display = asset_display & vbNewLine & "SECU - " & the_asset
            ElseIf ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                current_asset_panel = TRUE
                cars_panels = cars_panels + 1
                If DateDiff("d", ASSETS_ARRAY(update_date, the_asset), date) = 0 Then
                    ASSETS_ARRAY(ast_verif_date, the_asset) = doc_date_stamp
                    actions_taken =  actions_taken & "Updated CARS " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & ", "
                End If
                asset_display = asset_display & vbNewLine & "CARS - " & the_asset
            Else
                asset_display = asset_display & vbNewLine & ASSETS_ARRAY(ast_panel, the_asset) & " - " & the_asset
            End If
        Next

        'MsgBox asset_display

        dlg_len = 90

        If acct_panels > 0 Then dlg_len = dlg_len + 10
        dlg_len = dlg_len + (20 * acct_panels)
        If secu_panels > 0 Then dlg_len = dlg_len + 10
        dlg_len = dlg_len + (20 * secu_panels)
        If cars_panels > 0 Then dlg_len = dlg_len + 10
        dlg_len = dlg_len + (20 * cars_panels)
        'MsgBox dlg_len
        If acct_panels = 0 AND secu_panels = 0 AND  cars_panels = 0 Then run_updater_checkbox = checked

		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
        y_pos = 55
        BeginDialog Dialog1, 0, 0, 390, dlg_len, "Asset Verification Detail for Case #" & MAXIS_case_number
          ' Text 10, 15, 95, 10, "Date the form was received:"
          ' EditBox 110, 10, 35, 15, asset_form_doc_date
          Text 10, 15, 50, 10, "Action Taken:"
          EditBox 60, 10, 325, 15, actions_taken
          Text 20, 30, 350, 20, "                  *** To include any of these asset panel information in the CASE/NOTE ***                                   Enter detail about the asset/verification in the 'Verification  detail' field next  to the  panel information."
          If acct_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current ACCT panel details."
              Text 230, y_pos, 120, 10, "Verif date:"
              Text 280, y_pos, 50, 10, "Note:"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                      Text 10, y_pos, 275, 10,  "* ACCT " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset) & " - Balance: $" & ASSETS_ARRAY(ast_balance, the_asset)
                      EditBox 230, y_pos - 5, 45, 15, ASSETS_ARRAY(ast_verif_date, the_asset)
                      EditBox 280, y_pos - 5, 105, 15, ASSETS_ARRAY(ast_note, the_asset)
                      y_pos = y_pos + 20
                  End If
              Next
              y_pos = y_pos - 5
          End If

          If secu_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current SECU panel details."
              Text 230, y_pos, 120, 10, "Verif date:"
              Text 280, y_pos, 50, 10, "Note:"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                      Text 10, y_pos, 275, 10, "* SECU " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset)
                      EditBox 230, y_pos - 5, 45, 15, ASSETS_ARRAY(ast_verif_date, the_asset)
                      EditBox 280, y_pos - 5, 105, 15, ASSETS_ARRAY(ast_note, the_asset)
                      y_pos = y_pos + 20
                  End If
              Next
              y_pos = y_pos - 5
          End If

          If cars_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current CARS panel details."
              Text 230, y_pos, 120, 10, "Verif date:"
              Text 280, y_pos, 50, 10, "Note:"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                      Text 10, y_pos, 275, 10, "* CARS " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset)
                      EditBox 230, y_pos - 5, 45, 15, ASSETS_ARRAY(ast_verif_date, the_asset)
                      EditBox 280, y_pos - 5, 105, 15, ASSETS_ARRAY(ast_note, the_asset)
                      y_pos = y_pos + 20
                  End If
              Next
              y_pos = y_pos - 5
          End If
          y_pos = y_pos + 5

          CheckBox 20, y_pos, 250, 10, "Check here to have the script update asset panels. (ACCT, SECU, CARS).", run_updater_checkbox
          Text 30, y_pos + 10, 200, 10, "*Panels update by the script will be case noted."
          ButtonGroup ButtonPressed
            OkButton 280, y_pos, 50, 15
            CancelButton 335, y_pos, 50, 15
        EndDialog


        Do
            Do
                err_msg = ""
                dialog Dialog1
                Call cancel_continue_confirmation(skip_asset)
                IF actions_taken = "" AND run_updater_checkbox = unchecked THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered
                If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
                If skip_asset= TRUE Then
                    err_msg = ""
                    asset_form_checkbox = unchecked
                End If
                If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
            Loop Until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(ast_verif_date, the_asset) <> "" Then ASSETS_ARRAY(cnote_panel, the_asset) = checked
        Next
    End If

End If

highest_asset = asset_counter
If asset_form_checkbox = checked Then
    end_msg = end_msg & vbNewLine & "Asset detail entered."
    If LTC_case = vbNo Then docs_rec = docs_rec & ", Personal Statement (DHS 6054)"
    If LTC_case = vbYes Then docs_rec = docs_rec & ", Asset documents"
    If run_updater_checkbox = checked Then
        MAXIS_footer_month = CM_mo
        MAXIS_footer_year = CM_yr

        Do
            Call back_to_SELF
            Call MAXIS_background_check

            found_the_panel = FALSE
            panel_found = FALSE
            update_panel_type = "NONE - I'm all done"
            snap_is_yes = FALSE
			'-------------------------------------------------------------------------------------------------DIALOG
			Dialog1 = "" 'Blanking out previous dialog detail
            'Dialog to chose the panel type'
            BeginDialog Dialog1, 0, 0, 176, 85, "Type of panel to update"
              DropListBox 15, 25, 155, 45, "NONE - I'm all done"+chr(9)+"Existing ACCT"+chr(9)+"New ACCT"+chr(9)+"Existing SECU"+chr(9)+"New SECU"+chr(9)+"Existing CARS"+chr(9)+"New CARS", update_panel_type
              EditBox 90, 45, 20, 15, MAXIS_footer_month
              EditBox 115, 45, 20, 15, MAXIS_footer_year
              ButtonGroup ButtonPressed
                OkButton 120, 65, 50, 15
              Text 10, 10, 125, 10, "What panelwould you like to update?"
              Text 15, 50, 65, 10, "Footer Month/Year"
            EndDialog

            Do
                Do
                    err_msg = ""
                    dialog Dialog1
                    cancel_confirmation
                    If update_panel_type = "Existing ACCT" AND acct_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known ACCT panels, cannot update an 'Existing ACCT' panel."
                    If update_panel_type = "Existing SECU" AND secu_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known SECU panels, cannot update an 'Existing SECU' panel."
                    If update_panel_type = "Existing CARS" AND cars_panels = 0 Then err_msg = err_msg & vbNewLine & "* There are no known CARS panels, cannot update an 'Existing CARS' panel."
                    If Err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
                Loop until err_msg = ""
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

            panel_type = right(update_panel_type, 4)
            skip_this_panel = FALSE

            If panel_type = "ACCT" Then
                If update_panel_type = "Existing ACCT" Then
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "ACCT")
                        EMReadScreen navigate_check, 4, 2, 44
                    Loop until navigate_check = "ACCT"
                    For each member in HH_member_array
                        Call write_value_and_transmit(member, 20, 76)

                        EMReadScreen acct_versions, 1, 2, 78
                        If acct_versions <> "0" Then
                            EMWriteScreen "01", 20, 79
                            transmit
                            Do
                                is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")

                                If is_this_the_panel = vbYes Then found_the_panel = TRUE

                                If found_the_panel = TRUE then
                                    current_member = member
                                    Exit Do
                                End If
                                transmit
                                'MsgBox asset_counter
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If found_the_panel = TRUE then Exit For
                    Next

                    EMReadScreen current_instance, 1, 2, 73
                    current_instance = "0" & current_instance
                    For the_asset  = 0 to UBound(ASSETS_ARRAY, 2)
                        'MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
                        If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
                            asset_counter = the_asset
                            If ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y" Then count_cash_checkbox = checked
                            If ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then count_snap_checkbox = checked
                            If ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y" Then count_hc_checkbox = checked
                            If ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y" Then count_grh_checkbox = checked
                            If ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y" Then count_ive_checkbox = checked
                            'MsgBox ASSETS_ARRAY(ast_own_ratio, asset_counter)
                            share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            Exit For
                        End If
                    Next

                ElseIf update_panel_type = "New ACCT" Then
                    ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                End If

                If share_ratio_num = "" Then share_ratio_num = "1"
                If share_ratio_denom = "" Then share_ratio_denom = "1"
                If LTC_case = vbNo AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"

                ASSETS_ARRAY(ast_verif_date, asset_counter) = doc_date_stamp
				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
                'Dialog to fill the ACCT panel
                BeginDialog Dialog1, 0, 0, 271, 235, "New ACCT panel for Case #" & MAXIS_case_number
                  DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
                  DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+"SV - Savings"+chr(9)+"CK - Checking"+chr(9)+"CE - Certificate of Deposit"+chr(9)+"MM - Money Market"+chr(9)+"DC - Debit Card"+chr(9)+"KO - Keogh Account"+chr(9)+"FT - Federal Thrift Savings Plan"+chr(9)+"SL - State and Local Govt Ret"+chr(9)+"RA - Employee Ret Annuities"+chr(9)+"NP - Non-Profit Employer Ret Plans"+chr(9)+"IR - Indiv Ret Acct"+chr(9)+"RH - Roth IRA"+chr(9)+"FR - Ret Plans for Certain Employees"+chr(9)+"CT - Corp Ret Trust (before 1959)"+chr(9)+"RT - Other Ret Fund"+chr(9)+"QT - Qualified Tuition (529)"+chr(9)+"CA - Coverdell SV (530)"+chr(9)+"OE - Other Educationsal"+chr(9)+"OT - Other Account Type", ASSETS_ARRAY(ast_type, asset_counter)
                  EditBox 75, 50, 105, 15, ASSETS_ARRAY(ast_number, asset_counter)
                  EditBox 75, 70, 105, 15, ASSETS_ARRAY(ast_location, asset_counter)
                  EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_balance, asset_counter)
                  EditBox 160, 90, 50, 15, ASSETS_ARRAY(ast_bal_date, asset_counter)
                  DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_verif, asset_counter)
                  If LTC_case = vbYes Then EditBox 75, 130, 50, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
                  CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
                  CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
                  CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
                  CheckBox 230, 70, 30, 10, "GRH", count_grh_checkbox
                  CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
                  EditBox 75, 165, 50, 15, ASSETS_ARRAY(ast_wdrw_penlty, asset_counter)
                  DropListBox 75, 185, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_wthdr_verif, asset_counter)
                  EditBox 215, 125, 15, 15, share_ratio_num
                  EditBox 240, 125, 15, 15, share_ratio_denom
                  ComboBox 170, 160, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_one, asset_counter)
                  ComboBox 170, 175, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_two, asset_counter)
                  ComboBox 170, 190, 90, 45, client_dropdown_CB, ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
                  EditBox 75, 210, 50, 15, ASSETS_ARRAY(ast_next_inrst_date, asset_counter)
                  ButtonGroup ButtonPressed
                    OkButton 160, 215, 50, 15
                    CancelButton 215, 215, 50, 15
                  Text 10, 15, 60, 10, "Owner of Account:"
                  Text 20, 35, 50, 10, "Account Type:"
                  Text 15, 55, 60, 10, "Account Number:"
                  Text 10, 75, 60, 10, "Account Location:"
                  Text 40, 95, 30, 10, "Balance:"
                  Text 130, 95, 25, 10, "As of:"
                  Text 30, 115, 40, 10, "Verification:"
                  GroupBox 225, 10, 40, 90, "Count:"
                  GroupBox 20, 150, 140, 55, "Withdrawal Penalty"
                  Text 40, 170, 30, 10, "Amount:"
                  Text 30, 190, 40, 10, "Verification:"
                  If LTC_case = vbYes Then Text 35, 135, 35, 10, "Verif Date:"
                  GroupBox 165, 110, 100, 100, "Additional Owner(s)"
                  Text 170, 130, 40, 10, "Share Ratio:"
                  Text 170, 145, 50, 10, "Other owners:"
                  Text 5, 215, 65, 10, "Next Interest Date:"
                  Text 235, 125, 5, 10, "/"
                EndDialog

                Do
                    Do
                        err_msg = ""
                        dialog Dialog1
                        Call cancel_continue_confirmation(skip_this_panel)
                        ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(ASSETS_ARRAY(ast_wdrw_penlty, asset_counter))
                        ASSETS_ARRAY(ast_number, asset_counter) = trim(ASSETS_ARRAY(ast_number, asset_counter))
                        ASSETS_ARRAY(ast_location, asset_counter) = trim(ASSETS_ARRAY(ast_location, asset_counter))
                        ASSETS_ARRAY(ast_next_inrst_date, asset_counter) = trim(ASSETS_ARRAY(ast_next_inrst_date, asset_counter))
                        share_ratio_num = trim(share_ratio_num)
                        share_ratio_denom = trim(share_ratio_denom)
                        If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the bank account. The person must be listed in the household to have a new ACCT panel added."
                        If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of account this is."
                        If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Select the verification source for this account."
                        If ASSETS_ARRAY(ast_number, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_number, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The account number is too long."
                        If ASSETS_ARRAY(ast_location, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_location, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The location name is too long."
                        If IsNumeric(ASSETS_ARRAY(ast_balance, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance should be entered as a number."
                        If ASSETS_ARRAY(ast_bal_date, asset_counter) <> "" AND IsDate(ASSETS_ARRAY(ast_bal_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance effective date should be entered as a date."
                        If IsNumeric(share_ratio_num) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_num > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
                        End If
                        If IsNumeric(share_ratio_denom) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_denom > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
                        End If
                        If ASSETS_ARRAY(ast_next_inrst_date, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_next_inrst_date, asset_counter)) <> 5 Then err_msg = err_msg & vbNewLine & "* The next interest date should be entered in the format MM/YY."

                        If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
                        Else
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
                            If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
                        End If
                        If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
                        If skip_this_panel = TRUE Then
                            err_msg = ""
                            If update_panel_type = "New ACCT" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
                        End If
                        If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                    Loop until err_msg = ""
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = FALSE

                If skip_this_panel = FALSE Then
                    ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)
                    If count_cash_checkbox = checked Then ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y"
                    If count_snap_checkbox = checked Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    If count_hc_checkbox = checked Then ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y"
                    If count_grh_checkbox = checked Then ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y"
                    If count_ive_checkbox = checked Then ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y"
                    If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
                    If share_ratio_denom = "1" Then
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
                    Else
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
                        ASSETS_ARRAY(ast_share_note, asset_counter) = "ACCT is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
                    End If
                    ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
                    If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = ""
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "ACCT")
                        EMReadScreen navigate_check, 4, 2, 44
                    Loop until navigate_check = "ACCT"
                    EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
                    If update_panel_type = "Existing ACCT" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
                    transmit
                    If update_panel_type = "New ACCT" Then
                        EMWriteScreen "NN", 20, 79
                        transmit
                    End If
                    If update_panel_type = "Existing ACCT" Then PF9
                    ASSETS_ARRAY(cnote_panel, asset_counter) = checked
                    ASSETS_ARRAY(ast_panel, asset_counter) = "ACCT"
                    Call update_ACCT_panel_from_dialog
                    actions_taken =  actions_taken & "Updated ACCT " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "
                    If update_panel_type = "New ACCT" Then
                        EMReadScreen the_instance, 1, 2, 73
                        ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
                    End If
                    transmit
                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            panel_found = FALSE
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44
                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")
                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""
                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If
                        Call update_ACCT_panel_from_dialog
                        transmit
                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If
                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            panel_found = FALSE
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44
                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")
                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""

                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If

                        Call update_ACCT_panel_from_dialog
                        transmit

                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If

                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            panel_found = FALSE
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44
                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")
                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""
                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If
                        Call update_ACCT_panel_from_dialog
                        transmit
                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If
                End If
                if update_panel_type = "New ACCT" Then asset_counter = asset_counter + 1
                if update_panel_type = "Existing ACCT" Then asset_counter = highest_asset
            ElseIf panel_type = "SECU" Then
                If update_panel_type = "Existing SECU" Then
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "SECU")
                        EMReadScreen navigate_check, 4, 2, 45
                    Loop until navigate_check = "SECU"
                    For each member in HH_member_array
                        Call write_value_and_transmit(member, 20, 76)
                        EMReadScreen secu_versions, 1, 2, 78
                        If secu_versions <> "0" Then
                            EMWriteScreen "01", 20, 79
                            transmit
                            Do
                                is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")
                                If is_this_the_panel = vbYes Then found_the_panel = TRUE
                                If found_the_panel = TRUE then
                                    current_member = member
                                    Exit Do
                                End If
                                transmit
                                'MsgBox asset_counter
                                EMReadScreen reached_last_SECU_panel, 13, 24, 2
                            Loop until reached_last_SECU_panel = "ENTER A VALID"
                        End If
                        If found_the_panel = TRUE then Exit For
                    Next

                    EMReadScreen current_instance, 1, 2, 73
                    current_instance = "0" & current_instance
                    For the_asset  = 0 to UBound(ASSETS_ARRAY, 2)
                        'MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
                        If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
                            asset_counter = the_asset
                            If ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y" Then count_cash_checkbox = checked
                            If ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then count_snap_checkbox = checked
                            If ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y" Then count_hc_checkbox = checked
                            If ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y" Then count_grh_checkbox = checked
                            If ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y" Then count_ive_checkbox = checked
                            'MsgBox ASSETS_ARRAY(ast_own_ratio, asset_counter)
                            share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            Exit For
                        End If
                    Next

                Else update_panel_type = "New SECU"
                    ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                End If
                If share_ratio_num = "" Then share_ratio_num = "1"
                If share_ratio_denom = "" Then share_ratio_denom = "1"
                If LTC_case = vbNo AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"
                ASSETS_ARRAY(ast_verif_date, asset_counter) = doc_date_stamp

				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
                ' MsgBox ASSETS_ARRAY(ast_type, asset_counter)
                'Dialog to fill the SECU panel
                BeginDialog Dialog1, 0, 0, 271, 235, "New SECU panel for Case #" & MAXIS_case_number
                  DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
                  DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+"LI - Life Insurance"+chr(9)+"ST - Stocks"+chr(9)+"BO - Bonds"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"MO - Mortgage Note"+chr(9)+"AN - Annuity"+chr(9)+"OT - Other", ASSETS_ARRAY(ast_type, asset_counter)
                  EditBox 75, 50, 105, 15, ASSETS_ARRAY(ast_number, asset_counter)
                  EditBox 75, 70, 105, 15, ASSETS_ARRAY(ast_location, asset_counter)
                  EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_csv, asset_counter)
                  EditBox 160, 90, 50, 15, ASSETS_ARRAY(ast_bal_date, asset_counter)
                  DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Agency Form"+chr(9)+"2 - Source Doc"+chr(9)+"3 - Phone Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prov", ASSETS_ARRAY(ast_verif, asset_counter)
                  If LTC_case = vbYes Then EditBox 95, 130, 60, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
                  EditBox 95, 150, 60, 15, ASSETS_ARRAY(ast_face_value, asset_counter)
                  CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
                  CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
                  CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
                  CheckBox 230, 70, 30, 10, "GRH", count_grh_checkbox
                  CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
                  EditBox 75, 190, 50, 15, ASSETS_ARRAY(ast_wdrw_penlty, asset_counter)
                  DropListBox 75, 210, 80, 45, "Select..."+chr(9)+"1 - Agency Form"+chr(9)+"2 - Source Doc"+chr(9)+"3 - Phone Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prov", ASSETS_ARRAY(ast_wthdr_verif, asset_counter)
                  EditBox 215, 125, 15, 15, share_ratio_num
                  EditBox 240, 125, 15, 15, share_ratio_denom
                  ComboBox 170, 160, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_one, asset_counter)
                  ComboBox 170, 175, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_two, asset_counter)
                  ComboBox 170, 190, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
                  ButtonGroup ButtonPressed
                    OkButton 160, 215, 50, 15
                    CancelButton 215, 215, 50, 15
                  Text 10, 15, 60, 10, "Owner of Security:"
                  Text 20, 35, 50, 10, "Security Type:"
                  Text 10, 55, 60, 10, "Security Number:"
                  Text 15, 75, 55, 10, "Security Name:"
                  Text 10, 95, 60, 10, "Cash Value (CSV):"
                  Text 25, 115, 40, 10, "Verification:"
                  Text 130, 95, 25, 10, "As of:"
                  If LTC_case = vbYes Then Text 50, 135, 35, 10, "Verif Date:"
                  GroupBox 225, 10, 40, 90, "Count:"
                  Text 10, 155, 75, 10, "Face Value of Life Ins:"
                  GroupBox 20, 175, 140, 55, "Withdrawal Penalty"
                  Text 40, 195, 30, 10, "Amount:"
                  Text 30, 215, 40, 10, "Verification:"
                  GroupBox 165, 110, 100, 100, "Additional Owner(s)"
                  Text 170, 130, 40, 10, "Share Ratio:"
                  Text 170, 145, 50, 10, "Other owners:"
                  Text 235, 125, 5, 10, "/"
                EndDialog

                Do
                    Do
                        err_msg = ""
                        dialog Dialog1
                        Call cancel_continue_confirmation(skip_this_panel)
                        ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = trim(ASSETS_ARRAY(ast_wdrw_penlty, asset_counter))
                        ASSETS_ARRAY(ast_number, asset_counter) = trim(ASSETS_ARRAY(ast_number, asset_counter))
                        ASSETS_ARRAY(ast_location, asset_counter) = trim(ASSETS_ARRAY(ast_location, asset_counter))
                        ASSETS_ARRAY(ast_face_value, asset_counter) = trim(ASSETS_ARRAY(ast_face_value, asset_counter))
                        share_ratio_num = trim(share_ratio_num)
                        share_ratio_denom = trim(share_ratio_denom)
                        If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the security. The person must be listed in the household to have a new SECU panel added."
                        If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of security this is."
                        If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Select the verification source for this account."
                        If ASSETS_ARRAY(ast_number, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_number, asset_counter)) > 12 Then err_msg = err_msg & vbNewLine & "* The account number is too long."
                        If ASSETS_ARRAY(ast_location, asset_counter) <> "" AND len(ASSETS_ARRAY(ast_location, asset_counter)) > 20 Then err_msg = err_msg & vbNewLine & "* The location name is too long."
                        If IsNumeric(ASSETS_ARRAY(ast_csv, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance should be entered as a number."
                        If ASSETS_ARRAY(ast_bal_date, asset_counter) <> "" AND IsDate(ASSETS_ARRAY(ast_bal_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The balance effective date should be entered as a date."
                        If left(ASSETS_ARRAY(ast_type, asset_counter), 2) = "LI" Then
                            If ASSETS_ARRAY(ast_face_value, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* A life insurance policy requires a face value."
                            If count_snap_checkbox = checked Then
                                count_snap_checkbox = unchecked
                            End If
                        Else
                            If ASSETS_ARRAY(ast_face_value, asset_counter) <> "" Then err_msg = err_msg & vbNewLine & "* A face value amount can only be entered for a Life Insurance security."
                        End If
                        If IsNumeric(share_ratio_num) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_num > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
                        End If
                        If IsNumeric(share_ratio_denom) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_denom > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
                        End If
                        If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
                        Else
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
                            If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
                        End If
                        If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
                        If skip_this_panel = TRUE Then
                            err_msg = ""
                            If update_panel_type = "New SECU" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
                        End If
                        If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                    Loop until err_msg = ""
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = FALSE

                If skip_this_panel = FALSE Then
                    ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)
                    If count_cash_checkbox = checked Then ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y"
                    If count_snap_checkbox = checked Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    If count_hc_checkbox = checked Then ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y"
                    If count_grh_checkbox = checked Then ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y"
                    If count_ive_checkbox = checked Then ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y"
                    If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
                    If share_ratio_denom = "1" Then
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
                    Else
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
                        ASSETS_ARRAY(ast_share_note, asset_counter) = "SECU is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
                    End If
                    ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
                    If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = ""
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "SECU")
                        EMReadScreen navigate_check, 4, 2, 45
                    Loop until navigate_check = "SECU"
                    EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
                    If update_panel_type = "Existing SECU" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
                    transmit
                    If update_panel_type = "New SECU" Then
                        EMWriteScreen "NN", 20, 79
                        transmit
                    End If
                    If update_panel_type = "Existing SECU" Then PF9
                    ASSETS_ARRAY(cnote_panel, asset_counter) = checked
                    ASSETS_ARRAY(ast_panel, asset_counter) = "SECU"
                    Call update_SECU_panel_from_dialog
                    actions_taken =  actions_taken & "Updated SECU " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "

                    If update_panel_type = "New SECU" Then
                        EMReadScreen the_instance, 1, 2, 73
                        ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
                    End If
                    transmit

                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_one, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        panel_found = FALSE
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44

                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")

                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""

                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If

                        Call update_SECU_panel_from_dialog
                        transmit

                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If

                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_two, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            panel_found = FALSE
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44

                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")

                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""

                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If

                        Call update_SECU_panel_from_dialog
                        transmit

                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If

                    If IsNumeric(left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)) = True Then
                        ASSETS_ARRAY(ast_share_note, asset_counter) = ASSETS_ARRAY(ast_share_note, asset_counter) & " Also owned by M" & left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2)
                        EMWriteScreen left(ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter), 2), 20, 76
                        EMWriteScreen "01", 20, 79
                        transmit
                        EMReadScreen total_panels, 1, 2, 78
                        If total_panels = "0" Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                            panel_found = TRUE
                        Else
                            panel_found = FALSE
                            Do
                                EMReadScreen this_account_type, 2, 6, 44
                                EMReadScreen this_account_number, 20, 7, 44
                                EMReadScreen this_account_location, 20, 8, 44
                                this_account_number = replace(this_account_number, "_", "")
                                this_account_location = replace(this_account_location, "_", "")

                                If this_account_type = left(ASSETS_ARRAY(ast_type, asset_counter), 2) AND this_account_number = ASSETS_ARRAY(ast_number, asset_counter) AND this_account_location = ASSETS_ARRAY(ast_location, asset_counter) Then
                                    PF9
                                    panel_found = TRUE
                                    Exit Do
                                End If
                                transmit
                                EMReadScreen reached_last_ACCT_panel, 13, 24, 2
                            Loop until reached_last_ACCT_panel = "ENTER A VALID"
                        End If
                        If panel_found = FALSE Then
                            EMWriteScreen "NN", 20, 79
                            transmit
                        End If
                        panel_found = ""

                        IF ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then
                            snap_is_yes = TRUE
                            ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "N"
                        End If

                        Call update_ACCT_panel_from_dialog
                        transmit

                        If snap_is_yes = TRUE Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                    End If

                End If
                if update_panel_type = "New SECU" Then asset_counter = asset_counter + 1
                if update_panel_type = "Existing SECU" Then asset_counter = highest_asset
            ElseIf panel_type = "CARS" Then
                If update_panel_type = "Existing CARS" Then
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "CARS")
                        EMReadScreen navigate_check, 4, 2, 44
                    Loop until navigate_check = "CARS"
                    For each member in HH_member_array
                        Call write_value_and_transmit(member, 20, 76)

                        EMReadScreen cars_versions, 1, 2, 78
                        If acct_versions <> "0" Then
                            EMWriteScreen "01", 20, 79
                            transmit
                            Do
                                is_this_the_panel = MsgBox("Is this the panel you wish to update?", vbQuestion + vbYesNo, "Update this panel?")

                                If is_this_the_panel = vbYes Then found_the_panel = TRUE

                                If found_the_panel = TRUE then
                                    current_member = member
                                    Exit Do
                                End If
                                transmit
                                'MsgBox asset_counter
                                EMReadScreen reached_last_CARS_panel, 13, 24, 2
                            Loop until reached_last_CARS_panel = "ENTER A VALID"
                        End If
                        If found_the_panel = TRUE then Exit For
                    Next

                    EMReadScreen current_instance, 1, 2, 73
                    current_instance = "0" & current_instance
                    For the_asset  = 0 to UBound(ASSETS_ARRAY, 2)
                        'MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
                        If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" AND current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
                            asset_counter = the_asset

                            'MsgBox ASSETS_ARRAY(ast_own_ratio, asset_counter)
                            share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            Exit For
                        End If
                    Next

                ElseIf update_panel_type = "New CARS" Then
                    ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                End If

                If share_ratio_num = "" Then share_ratio_num = "1"
                If share_ratio_denom = "" Then share_ratio_denom = "1"
                If LTC_case = vbNo AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "5 - Other Document"
                ASSETS_ARRAY(ast_verif_date, asset_counter) = doc_date_stamp

				'-------------------------------------------------------------------------------------------------DIALOG
				Dialog1 = "" 'Blanking out previous dialog detail
                'Dialog to fill the CARS panel.
                BeginDialog Dialog1, 0, 0, 270, 255, "New CARS panel for Case # & MAXIS_case_number" & MAXIS_case_number
                  DropListBox 75, 10, 135, 45, client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
                  DropListBox 75, 30, 90, 45, "Select..."+chr(9)+"1 - Car"+chr(9)+"2 - Truck"+chr(9)+"3 - Van"+chr(9)+"4 - Camper"+chr(9)+"5 - Motorcycle"+chr(9)+"6 - Trailer"+chr(9)+"7 - Other", ASSETS_ARRAY(ast_type, asset_counter)
                  EditBox 220, 30, 40, 15, ASSETS_ARRAY(ast_year, asset_counter)
                  ComboBox 75, 50, 185, 45, "Type or Select"+chr(9)+"Acura"+chr(9)+"Audi"+chr(9)+"BMW"+chr(9)+"Buick"+chr(9)+"Cadillac"+chr(9)+"Chevrolet"+chr(9)+"Chrysler"+chr(9)+"Dodge"+chr(9)+"Ford"+chr(9)+"GMC"+chr(9)+"Honda"+chr(9)+"Hummer"+chr(9)+"Hyundai"+chr(9)+"Infiniti"+chr(9)+"Isuzu"+chr(9)+"Jeep"+chr(9)+"Kia"+chr(9)+"Lincoln"+chr(9)+"Mazda"+chr(9)+"Mercedes-Benz"+chr(9)+"Mercury"+chr(9)+"Mitsubishi"+chr(9)+"Nissan"+chr(9)+"Oldsmobile"+chr(9)+"Plymouth"+chr(9)+"Pontiac"+chr(9)+"Saab"+chr(9)+"Saturn"+chr(9)+"Scion"+chr(9)+"Subaru"+chr(9)+"Suzuki"+chr(9)+"Toyota"+chr(9)+"Volkswagen"+chr(9)+"Volvo", ASSETS_ARRAY(ast_make, asset_counter)
                  EditBox 75, 70, 185, 15, ASSETS_ARRAY(ast_model, asset_counter)
                  EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_trd_in, asset_counter)
                  DropListBox 165, 90, 95, 45, "Select..."+chr(9)+"1 - NADA"+chr(9)+"2 - Appraisal Value"+chr(9)+"3 - Client Stmt"+chr(9)+"4 - Other Document", ASSETS_ARRAY(ast_value_srce, asset_counter)
                  DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Title"+chr(9)+"2 - License Reg"+chr(9)+"3 - DMV"+chr(9)+"4 - Purchase Agmt"+chr(9)+"5 - Other Document"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_verif, asset_counter)
                  If LTC_case = vbYes Then EditBox 210, 110, 50, 15, ASSETS_ARRAY(ast_verif_date, asset_counter)
                  DropListBox 75, 130, 60, 45, "No"+chr(9)+"Yes", ASSETS_ARRAY(ast_hc_benefit, asset_counter)
                  EditBox 75, 165, 50, 15, ASSETS_ARRAY(ast_amt_owed, asset_counter)
                  EditBox 75, 185, 50, 15, ASSETS_ARRAY(ast_owed_date, asset_counter)
                  DropListBox 75, 210, 80, 45, "Select..."+chr(9)+"1 - Bank/Lending Inst Stmt"+chr(9)+"2 - Private Lender Stmt"+chr(9)+"3 - Other Document"+chr(9)+"4 - Pend Out State Verif"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_owe_verif, asset_counter)
                  EditBox 215, 145, 15, 15, share_ratio_num
                  EditBox 240, 145, 15, 15, share_ratio_denom
                  ComboBox 170, 180, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_oner, asset_counter)
                  ComboBox 170, 195, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_twor, asset_counter)
                  ComboBox 170, 210, 90, 45, "Type or Select", ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
                  ButtonGroup ButtonPressed
                    OkButton 170, 235, 45, 15
                    CancelButton 220, 235, 45, 15
                  Text 10, 15, 60, 10, "Owner of Vehicle:"
                  Text 20, 35, 50, 10, "Vehicle Type:"
                  Text 170, 35, 45, 10, "Vehicle Year:"
                  Text 20, 55, 50, 10, "Vehicle Make:"
                  Text 20, 75, 50, 10, "Vehicle Model:"
                  Text 20, 95, 50, 10, "Trade In Value:"
                  Text 130, 95, 25, 10, "Source:"
                  Text 25, 115, 40, 10, "Verification:"
                  If LTC_case = vbYes Then Text 170, 115, 40, 10, "Verif Date:"
                  Text 15, 135, 50, 10, "HC Clt Benefit:"
                  GroupBox 20, 150, 140, 80, "Amount Owed on vehicle"
                  Text 40, 170, 30, 10, "Amount:"
                  Text 45, 190, 20, 10, "As of:"
                  Text 30, 210, 40, 10, "Verification:"
                  GroupBox 165, 130, 100, 100, "Additional Owner(s)"
                  Text 170, 150, 40, 10, "Share Ratio:"
                  Text 170, 165, 50, 10, "Other owners:"
                  Text 235, 145, 5, 10, "/"
                EndDialog

                Do
                    Do
                        err_msg = ""
                        dialog Dialog1
                        Call cancel_continue_confirmation(skip_this_panel)
                        ASSETS_ARRAY(ast_year, asset_counter) = trim(ASSETS_ARRAY(ast_year, asset_counter))
                        ASSETS_ARRAY(ast_make, asset_counter) = trim(ASSETS_ARRAY(ast_make, asset_counter))
                        ASSETS_ARRAY(ast_model, asset_counter) = trim(ASSETS_ARRAY(ast_model, asset_counter))
                        ASSETS_ARRAY(ast_trd_in, asset_counter) = trim(ASSETS_ARRAY(ast_trd_in, asset_counter))
                        share_ratio_num = trim(share_ratio_num)
                        share_ratio_denom = trim(share_ratio_denom)
                        If ASSETS_ARRAY(ast_owner, asset_counter) = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the owner of the vehicle. The person must be listed in the household to have a new SECU panel added."
                        If ASSETS_ARRAY(ast_type, asset_counter) = "Select ..." Then err_msg = err_msg & vbNewLine & "* Indicate the type of vehicle this is."
                        If ASSETS_ARRAY(ast_year, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the year of the vehicle."
                        If len(ASSETS_ARRAY(ast_year, asset_counter)) <> 4 Then err_msg = err_msg & vbNewLine & "* The year of the vehicle needs to be in the format YYYY."
                        If ASSETS_ARRAY(ast_make, asset_counter) = "Type or Select" OR ASSETS_ARRAY(ast_make, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the make of the vehicle."
                        If ASSETS_ARRAY(ast_model, asset_counter) = "" Then err_msg = err_msg & vbNewLine & "* Enter the model of the vehicle."
                        If IsNumeric(ASSETS_ARRAY(ast_trd_in, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The trade in value needs to be entered as a number."
                        If ASSETS_ARRAY(ast_value_srce, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate from where the value was determined."
                        If ASSETS_ARRAY(ast_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Enter the verification of the vehicle."

                        If ASSETS_ARRAY(ast_amt_owed, asset_counter) <> "" Then
                            If IsNumeric(ASSETS_ARRAY(ast_amt_owed, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* The owed amount needs to be entered as a number."
                            If ASSETS_ARRAY(ast_owe_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* Enter the verification of the amount that owed."
                            If IsDate(ASSETS_ARRAY(ast_owed_date, asset_counter)) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for the effective date of the owed amount."
                        End If

                        If IsNumeric(share_ratio_num) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_num > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio top number must be 9 or lower"
                        End If
                        If IsNumeric(share_ratio_denom) = FALSE Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio must be entered in numerals."
                        ElseIf share_ratio_denom > 9 Then
                            err_msg = err_msg & vbNewLine & "* The Share Ratio bottom number must be 9 or lower"
                        End If

                        If ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0.00" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "0" OR ASSETS_ARRAY(ast_wdrw_penlty, asset_counter) = "" Then
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "N"
                        Else
                            ASSETS_ARRAY(ast_wthdr_YN, asset_counter) = "Y"
                            If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then err_msg = err_msg & vbNewLine & "* If there is a withdraw penalty amount listed, this amount needs a verification selected."
                        End If
                        If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
                        If skip_this_panel = TRUE Then
                            err_msg = ""
                            If update_panel_type = "New SECU" Then ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter - 1)
                        End If

                        If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                    Loop until err_msg = ""
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = FALSE

                If skip_this_panel = FALSE Then
                    ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)

                    If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
                    If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
                    If ASSETS_ARRAY(ast_owe_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_owe_verif, asset_counter) = ""
                    ASSETS_ARRAY(ast_loan_value, asset_counter) = .9 * ASSETS_ARRAY(ast_trd_in, asset_counter)
                    ASSETS_ARRAY(ast_loan_value, asset_counter) = round(ASSETS_ARRAY(ast_loan_value, asset_counter))
                    If share_ratio_denom = "1" Then
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
                    Else
                        ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "Y"
                        ASSETS_ARRAY(ast_share_note, asset_counter) = "CARS is shared. M" & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " owns " & share_ratio_num & "/" & share_ratio_denom & "."
                    End If
                    ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
                    If ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "Yes" Then ASSETS_ARRAY(ast_hc_benefit, asset_counter)  = "Y"
                    If ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "No" Then ASSETS_ARRAY(ast_hc_benefit, asset_counter) = "N"
                    Do
                        Call navigate_to_MAXIS_screen("STAT", "CARS")
                        EMReadScreen navigate_check, 4, 2, 44
                    Loop until navigate_check = "CARS"
                    EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
                    If update_panel_type = "Existing CARS" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
                    transmit
                    If update_panel_type = "New CARS" Then
                        EMWriteScreen "NN", 20, 79
                        transmit
                    End If
                    If update_panel_type = "Existing CARS" Then PF9

                    ASSETS_ARRAY(cnote_panel, asset_counter) = checked
                    ASSETS_ARRAY(ast_panel, asset_counter) = "CARS"

                    Call update_CARS_panel_from_dialog

                    actions_taken =  actions_taken & "Updated CARS " & ASSETS_ARRAY(ast_ref_nbr, asset_counter) & " " & ASSETS_ARRAY(ast_instance, asset_counter) & ", "


                    If update_panel_type = "New CARS" Then
                        EMReadScreen the_instance, 1, 2, 73
                        ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
                    End If
                    transmit


                End If

                if update_panel_type = "New CARS" Then asset_counter = asset_counter + 1
                if update_panel_type = "Existing CARS" Then asset_counter = highest_asset
            End If
            highest_asset = asset_counter

        Loop until panel_type = "done"
    End If

End If

If arep_form_checkbox = checked Then

    Call navigate_to_MAXIS_screen("STAT", "AREP")

    update_AREP_panel_checkbox = checked
    AREP_recvd_date = doc_date_stamp

    EMReadScreen arep_name, 37, 4, 32
    arep_name = replace(arep_name, "_", "")
    If arep_name <> "" Then
        EMReadScreen arep_street_one, 22, 5, 32
        EMReadScreen arep_street_two, 22, 6, 32
        EMReadScreen arep_city, 15, 7, 32
        EMReadScreen arep_state, 2, 7, 55
        EMReadScreen arep_zip, 5, 7, 64

        arep_street_one = replace(arep_street_one, "_", "")
        arep_street_two = replace(arep_street_two, "_", "")
        arep_street = arep_street_one & " " & arep_street_two
        arep_street = trim( arep_street)
        arep_city = replace(arep_city, "_", "")
        arep_state = replace(arep_state, "_", "")
        arep_zip = replace(arep_zip, "_", "")

        EMReadScreen arep_phone_one, 14, 8, 34
        EMReadScreen arep_ext_one, 3, 8, 55
        EMReadScreen arep_phone_two, 14, 9, 34
        EMReadScreen arep_ext_two, 3, 8, 55

        arep_phone_one = replace(arep_phone_one, ")", "")
        arep_phone_one = replace(arep_phone_one, "  ", "-")
        arep_phone_one = replace(arep_phone_one, " ", "-")
        If arep_phone_one = "___-___-____" Then arep_phone_one = ""

        arep_phone_two = replace(arep_phone_two, ")", "")
        arep_phone_two = replace(arep_phone_two, "  ", "-")
        arep_phone_two = replace(arep_phone_two, " ", "-")
        If arep_phone_two = "___-___-____" Then arep_phone_two = ""

        arep_ext_one = replace(arep_ext_one, "_", "")
        arep_ext_two = replace(arep_ext_two, "_", "")

        EMReadScreen forms_to_arep, 1, 10, 45
        EMReadScreen mmis_mail_to_arep, 1, 10, 77

        If forms_to_arep = "Y" Then forms_to_arep_checkbox = checked
        If mmis_mail_to_arep = "Y" Then mmis_mail_to_arep_checkbox = checked

        update_AREP_panel_checkbox = unchecked
    End If
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 396, 210, "AREP for Case #  & MAXIS_case_number"
      EditBox 40, 20, 215, 15, arep_name
      EditBox 40, 40, 215, 15, arep_street
      EditBox 40, 60, 85, 15, arep_city
      EditBox 160, 60, 20, 15, arep_state
      EditBox 215, 60, 40, 15, arep_zip
      EditBox 40, 80, 50, 15, arep_phone_one
      EditBox 110, 80, 20, 15, arep_ext_one
      EditBox 165, 80, 50, 15, arep_phone_two
      EditBox 235, 80, 20, 15, arep_ext_two
      CheckBox 15, 105, 60, 10, "Forms to AREP", forms_to_arep_checkbox
      CheckBox 90, 105, 75, 10, "MMIS Mail to AREP", mmis_mail_to_arep_checkbox
      CheckBox 15, 120, 185, 10, "Check here to have the script update the AREP Panel", update_AREP_panel_checkbox
      EditBox 110, 140, 50, 15, AREP_recvd_date
      CheckBox 10, 160, 75, 10, "ID on file for AREP?", AREP_ID_check
      CheckBox 10, 175, 215, 10, "TIKL to get new HC form 12 months after date form was signed?", TIKL_check
      EditBox 130, 190, 65, 15, arep_signature_date
      CheckBox 260, 175, 35, 10, "SNAP", SNAP_AREP_checkbox
      CheckBox 300, 175, 50, 10, "Health Care", HC_AREP_checkbox
      CheckBox 355, 175, 30, 10, "Cash", CASH_AREP_checkbox
      ButtonGroup ButtonPressed
        OkButton 285, 190, 50, 15
        CancelButton 340, 190, 50, 15
      GroupBox 5, 5, 255, 130, "Panel Information"
      Text 15, 25, 25, 10, "Name:"
      Text 15, 45, 25, 10, "Street:"
      Text 15, 65, 20, 10, "City:"
      Text 135, 65, 20, 10, "State:"
      Text 195, 65, 20, 10, "Zip:"
      Text 10, 85, 25, 10, "Phone:"
      Text 95, 85, 15, 10, "Ext."
      Text 140, 85, 25, 10, "Phone:"
      Text 220, 85, 15, 10, "Ext."
      Text 10, 145, 95, 10, "Date of AREP Form Received"
      Text 10, 195, 115, 10, "Date form was signed (MM/DD/YY):"
      Text 255, 160, 85, 10, "Programs Authorized for:"
      GroupBox 265, 5, 125, 150, "Specific FORM Received"
      CheckBox 275, 15, 115, 10, "AREP Req - MHCP - DHS-3437", dhs_3437_checkbox
      CheckBox 275, 35, 105, 10, "AREP Req - HC12729", HC_12729_checkbox
      CheckBox 275, 55, 100, 10, "SNAP AREP Choice - D405", D405_checkbox
      CheckBox 275, 75, 105, 10, "AREP on CAF", CAF_AREP_page_checkbox
      CheckBox 275, 95, 100, 10, "AREP on any HC App", HCAPP_AREP_checkbox
      CheckBox 275, 115, 75, 10, "Power of Attorney", power_of_attorney_checkbox
      Text 295, 25, 50, 10, "(HC)"
      Text 295, 45, 60, 10, "(Cash and SNAP)"
      Text 295, 65, 75, 10, "(SNAP and EBT Card)"
      Text 295, 85, 60, 10, "(Cash and SNAP)"
      Text 295, 105, 50, 10, "(HC)"
      Text 295, 125, 60, 10, "(HC, SNAP, Cash)"
      Text 270, 135, 110, 15, "Checking the FORM will indicate the programs in the CASE/NOTE"
    EndDialog

    Do
        Do
        	err_msg = ""
        	dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
        	cancel_confirmation
            cancel_continue_confirmation(skip_arep)
            If trim(arep_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the AREP's name."
            If update_AREP_panel_checkbox = checked Then
                If trim(arep_street) = "" OR trim(arep_city) = "" OR trim(arep_zip) = "" Then err_msg = err_msg & vbNewLine & "* Enter the street address of the AREP."
                If len(arep_name) > 37 Then err_msg = err_msg & vbNewLine & "* The AREP name is too long for MAXIS."
                If len(arep_street) > 44 Then err_msg = err_msg & vbNewLine & "* The AREP street is too long for MAXIS."
                If len(arep_city) > 15 Then err_msg = err_msg & vbNewLine & "* The AREP City is too long for MAXIS."
                If len(arep_state) > 2 Then err_msg = err_msg & vbNewLine & "* The AREP state is too long for MAXIS."
                If len(arep_zip) > 5 Then err_msg = err_msg & vbNewLine & "* The AREP zip is too long for MAXIS."
            End If
            If dhs_3437_checkbox = Checked Then HC_AREP_checkbox = checked
            If HC_12729_checkbox = checked Then
                SNAP_AREP_checkbox = checked
                CASH_AREP_checkbox = checked
            End If
            If D405_checkbox = checked Then SNAP_AREP_checkbox = checked
            If CAF_AREP_page_checkbox = checked Then
                SNAP_AREP_checkbox = checked
                CASH_AREP_checkbox = Checked
            End If
            If HCAPP_AREP_checkbox = checked Then HC_AREP_checkbox = checked
            If power_of_attorney_checkbox = checked Then
                SNAP_AREP_checkbox = checked
                CASH_AREP_checkbox = Checked
                HC_AREP_checkbox = checked
            End If
            If IsDate(AREP_recvd_date) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the form was received."
        	IF SNAP_AREP_checkbox <> checked AND HC_AREP_checkbox <> checked AND CASH_AREP_checkbox <> checked THEN err_msg = err_msg & vbNewLine &"* Select a program"
        	IF isdate(arep_signature_date) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was signed/valid from."
        	IF (TIKL_check = checked AND arep_signature_date = "") THEN err_msg = err_msg & vbNewLine & "* You have requested the script to TIKL based on the signature date but you did not enter the signature date."
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_arep = TRUE Then
                err_msg = ""
                arep_form_checkbox = unchecked
            End If
        	IF err_msg <> ""  AND left(err_msg,4) <> "LOOP" THEN MsgBox "Plese resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
End If

If arep_form_checkbox = checked Then
    end_msg = end_msg & vbNewLine & "AREP Information entered."
    'formatting programs into one variable to write in case note
    IF SNAP_AREP_checkbox = checked THEN AREP_programs = "SNAP"
    IF HC_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", HC"
    IF CASH_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", CASH"
    If left(AREP_programs, 1) = "," Then AREP_programs = right(AREP_programs, len(AREP_programs)-2)

    docs_rec = docs_rec & ", AREP Form"

    If update_AREP_panel_checkbox = checked Then
        Call MAXIS_background_check

        If IsDate(arep_signature_date) = TRUE Then
            Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, arep_signature_date)
        Else
            Call get_footer_month_from_date(MAXIS_footer_month, MAXIS_footer_year, doc_date_stamp)
        End If

        Call back_to_SELF
        Do
            Call navigate_to_MAXIS_screen("STAT", "AREP")
            EMReadScreen panel_check, 4, 2, 53
        Loop until panel_check = "AREP"

        EMReadScreen arep_version, 1, 2, 73
        If arep_version = "1" Then PF9
        If arep_version = "0" Then Call write_value_and_transmit("NN", 20, 79)

        'Writing to the panel
        EMWriteScreen "                                     ", 4, 32
        EMWriteScreen "                      ", 5, 32
        EMWriteScreen "                      ", 6, 32
        EMWriteScreen "               ", 7, 32
        EMWriteScreen "  ", 7, 55
        EMWriteScreen "     ", 7, 64

        EMWriteScreen arep_name, 4, 32
        arep_sreet = trim(arep_street)
        If len(arep_street) > 22 Then
            arep_street_one = ""
            arep_street_two = ""
            street_array = split(arep_street, " ")
            For each word in street_array
                If len(arep_street_one & word) > 22 Then
                    arep_street_two = arep_street_two & word & " "
                Else
                    arep_street_one = arep_street_one & word & " "
                End If
            Next
        Else
            arep_street_one = arep_street
        End If
        EMWriteScreen arep_street_one, 5, 32
        EMWriteScreen arep_street_two, 6, 32
        EMWriteScreen arep_city, 7, 32
        EMWriteScreen arep_state, 7, 55
        EMWriteScreen arep_zip, 7, 64
        EMWriteScreen "N", 5, 77

        If arep_phone_one <> "" Then
            write_phone_one = replace(arep_phone_one, "(", "")
            write_phone_one = replace(write_phone_one, ")", "")
            write_phone_one = replace(write_phone_one, "-", "")
            write_phone_one = trim(write_phone_one)

            EMWriteScreen left(write_phone_one, 3), 8, 34
            EMwriteScreen right(left(write_phone_one, 6), 3), 8, 40
            EMWriteScreen right(write_phone_one, 4), 8, 44

            If arep_ext_one = "" Then
                EMWriteScreen "   ", 8, 55
            Else
                EMWriteScreen arep_ext_one, 8, 55
            End If
        End If

        If arep_phone_two <> "" Then
            write_phone_two = replace(arep_phone_two, "(", "")
            write_phone_two = replace(write_phone_two, ")", "")
            write_phone_two = replace(write_phone_two, "-", "")
            write_phone_two = trim(write_phone_two)

            EMWriteScreen left(write_phone_two, 3), 8, 34
            EMwriteScreen right(left(write_phone_two, 6), 3), 8, 40
            EMWriteScreen right(write_phone_two, 4), 8, 44

            If arep_ext_two = "" Then
                EMWriteScreen "   ", 8, 55
            Else
                EMWriteScreen arep_ext_two, 8, 55
            End If
        End If

        If forms_to_arep_checkbox = checked Then EMWriteScreen "Y", 10, 45
        If forms_to_arep_checkbox = unchecked Then EMWriteScreen "N", 10, 45
        If mmis_mail_to_arep_checkbox = checked Then EMWriteScreen "Y", 10, 77
        If mmis_mail_to_arep_checkbox = unchecked Then EMWriteScreen "N", 10, 77

        transmit

    End If

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    If TIKL_check = checked then Call create_TIKL("Client's AREP release for HC is now 12 months old and no longer valid. Take appropriate action.", 365, arep_signature_date, False, TIKL_note_text)
End If

If LTC_case = vbNo Then

    If ADDR = "" AND SCHL = "" AND DISA = "" AND mof_form_checkbox = unchecked AND  JOBS = "" AND BUSI = "" AND evf_form_received_checkbox = unchecked AND UNEA = "" AND ACCT = "" AND asset_form_checkbox = unchecked AND SHEL = "" AND INSA = "" AND other_assets = "" AND arep_form_checkbox = unchecked AND other_verifs = "" AND notes = "" Then need_final_note = FALSE

End If

If LTC_case = vbYes Then

    If FACI = "" AND JOBS = "" AND BUSI_RBIC = "" AND evf_form_received_checkbox = unchecked AND UNEA = "" AND ACCT = "" AND asset_form_checkbox = unchecked AND SECU = "" AND CARS = "" AND REST = "" AND OTHR = "" AND SHEL = "" AND INSA = "" AND medical_expenses = "" AND arep_form_checkbox = unchecked AND veterans_info = "" AND other_verifs = "" AND notes = "" Then need_final_note = FALSE

End If

If mtaf_form_checkbox = checked Then
    MTAF_date = doc_date_stamp
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    Do
        Do
	      BeginDialog Dialog1, 0, 0, 186, 265, "MTAF dialog"
              EditBox 55, 5, 60, 15, MTAF_date
              DropListBox 55, 25, 60, 15, "Select one..."+chr(9)+"complete"+chr(9)+"incomplete", MTAF_status_dropdown
              EditBox 55, 45, 60, 15, MFIP_elig_date
              CheckBox 5, 65, 55, 10, "MTAF signed.", mtaf_signed_checkbox
              CheckBox 5, 80, 140, 10, "MFIP/financial orientation completed.", mfip_financial_orientation_checkbox
              CheckBox 5, 95, 150, 10, "Client exempt from cooperation with ES.", ES_exemption_checkbox
              CheckBox 5, 110, 180, 10, "Sent MFIP financial orientation DVD to participant(s).", MFIP_DVD_checkbox
              EditBox 55, 125, 60, 15, interview_date
              CheckBox 5, 145, 135, 10, "Rights and responsibilities explained.", RR_explained_checkbox
              ButtonGroup ButtonPressed
                OkButton 80, 245, 50, 15
                CancelButton 130, 245, 50, 15
              Text 5, 10, 40, 10, "MTAF date:"
              Text 5, 30, 45, 10, "MTAF status:"
              Text 5, 50, 50, 10, "MFIP elig date:"
              Text 5, 130, 50, 10, "Interview date:"
              GroupBox 5, 155, 175, 85, ""
              Text 15, 165, 155, 25, "*STOP WORK - Verification only necessary to verify income in the month of application/eligibility. (CM 0010.18.01)"
              Text 15, 200, 160, 35, "**SUBSIDY - Verification of housing subsidy is a mandatory verification for MFIP. STAT must be appropriately updated to ensure accurate approval of housing grant. (CM 0010.18.01)"
            EndDialog

            err_msg = ""

			DIALOG Dialog1
            cancel_continue_confirmation(skip_mtaf)
            If IsDate(MTAF_date) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the MTAF was received."
            If MTAF_status_dropdown = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate the status of the MTAF."
            'If  Then err_msg = err_msg & vbNewLine & "* "
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_mtaf = TRUE Then
                err_msg = ""
                mtaf_form_checkbox = unchecked
            End If

            If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox ("Please resolve to continue:" & vbNewLine & err_msg)
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE
End If

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
If mtaf_form_checkbox = checked Then
    BeginDialog Dialog1, 0, 0, 345, 350, "MTAF dialog"
      CheckBox 10, 10, 225, 10, "Check here if all other docs rec'vd are associated with this MTAF.", MTAF_note_only_checkbox
      EditBox 75, 50, 260, 15, ADDR_change
      EditBox 75, 70, 260, 15, HHcomp_change
      EditBox 75, 90, 260, 15, asset_change
      EditBox 105, 110, 230, 15, earned_income_change
      EditBox 105, 130, 230, 15, unearned_income_change
      EditBox 105, 150, 230, 15, shelter_costs_change
      EditBox 175, 170, 160, 15, subsidized_housing
      DropListBox 175, 190, 160, 15, "Select one..."+chr(9)+"Not subsidized"+chr(9)+"Verification provided"+chr(9)+"Verification pending", sub_housing_droplist
      EditBox 110, 205, 225, 15, child_adult_care_costs
      EditBox 110, 225, 225, 15, relationship_proof
      EditBox 175, 245, 160, 15, referred_to_OMB_PBEN
      EditBox 125, 265, 210, 15, elig_results_fiated
      EditBox 75, 285, 260, 15, other_notes
      EditBox 75, 305, 260, 15, verifications_needed
      ButtonGroup ButtonPressed
        OkButton 235, 330, 50, 15
        CancelButton 290, 330, 50, 15
      Text 20, 20, 225, 10, "Checking this box creates a MTAF CASE/NOTE ONLY."
      GroupBox 5, 35, 335, 290, "**Changes reported on MTAF**  (Complete boxes as applicable.)"
      Text 10, 55, 65, 10, "Address changes:"
      Text 10, 75, 65, 10, "HH comp changes:"
      Text 10, 95, 65, 10, "Assets changes:"
      Text 10, 115, 90, 10, "*Change in earned income:"
      Text 10, 135, 95, 10, "Change in unearned income:"
      Text 10, 155, 95, 10, "Change in shelter costs:"
      Text 10, 175, 165, 10, "Is housing subsidized? If so, what is the amount?"
      Text 80, 190, 95, 10, "**Subsidized housing status:"
      Text 10, 210, 85, 10, "Child or adult care costs:"
      Text 10, 230, 95, 10, "Proof of relationship on file:"
      Text 10, 250, 160, 10, "Client has been referred to apply for OMB/PBEN:"
      Text 10, 270, 115, 10, "Eligibility results fiated? If so, why:"
      Text 10, 290, 45, 10, "Other notes:"
      Text 10, 310, 65, 10, "Verifs needed:"
    EndDialog

    Do
        Do
            err_msg = ""
            dialog Dialog1
            cancel_continue_confirmation(skip_mtaf)
            If sub_housing_droplist = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate if housind is subsidized or not."
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_mtaf = TRUE Then
                err_msg = ""
                mtaf_form_checkbox = unchecked
            End If
            If err_msg <> ""  AND left(err_msg,4) <> "LOOP" Then MsgBox ("Please resolve to continue:" & vbNewLine & err_msg)
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE
End If

If mtaf_form_checkbox = checked Then
    end_msg = end_msg & vbNewLine & "MTAF Information entered."
    If MTAF_note_only_checkbox = checked Then
        need_final_note = FALSE
    End If
    'Takes script to a blank case note.
    Call start_a_blank_case_note

    'THE CASE NOTE===========================================================================
    CALL write_variable_in_CASE_NOTE("***MTAF Processed: " & MTAF_status_dropdown & "***")
    CALL write_bullet_and_variable_in_CASE_NOTE ("Date received", MTAF_date)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Date of eligibility", MFIP_elig_date)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Date of interview", interview_date)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Address change", ADDR_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Household composition change", HHcomp_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Change in assets", asset_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Change in earned income", earned_income_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Change in unearned income", unearned_income_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Change in shelter costs", shelter_costs_change)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Is housing subsidized? If so, what is the amount", subsidized_housing)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Subsidized housing status", sub_housing_droplist)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Child or adult care costs", child_adult_care_costs)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Proof of relationship on file", relationship_proof)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Referred to apply for OMB/PBEN", referred_to_OMB_PBEN)
    CALL write_bullet_and_variable_in_CASE_NOTE ("ELIG results fiated", elig_results_fiated)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Other notes", other_notes)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Verifications Needed", verifications_needed)
    IF MFIP_DVD_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent MFIP orientation DVD to participant(s).")
    If RR_explained_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Rights & responsibilities explained.")
    If mtaf_signed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MTAF was signed.")
    If mfip_financial_orientation_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* MFIP orientation information reviewed/completed.")
    If ES_exemption_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("* Client is exempt from cooperation with ES at this time.")
    CALL write_bullet_and_variable_in_CASE_NOTE ("MTAF Status", MTAF_status_dropdown)
    If MTAF_note_only_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("---")
        If HSR_scanner_checkbox = checked then
            Call write_variable_in_case_note("Docs Rec'd & scanned: " & docs_rec)
        else
            Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
        END IF
        call write_bullet_and_variable_in_case_note("Document date stamp", doc_date_stamp)
        If arep_form_checkbox = checked Then
            call write_variable_in_CASE_NOTE("* AREP FORM received on " & AREP_recvd_date & ". AREP: " & arep_name)
            If dhs_3437_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the DHS 3437 - MHCP AUTHORIZED REPRESENTATIVE REQUEST Form.")
            If HC_12729_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the HC 12729 - AUTHORIZED REPRESENTATIVE REQUEST Form.")
            If D405_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("  - AREP name on the SNAP AUTHORIZED REPRESENTATIVE CHOICE D405 Form.")
                Call write_variable_in_CASE_NOTE("  - AREP also authorixed to get and use EBT Card.")
            End If
            If CAF_AREP_page_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in the CAF.")
            If HCAPP_AREP_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in a Health Care Application.")
            If power_of_attorney_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP has Power of Attorney Designation.")
            If AREP_programs <> "" then call write_variable_in_CASE_NOTE("  - Programs Authorized for: " & AREP_programs)
            If arep_signature_date <> "" Then call write_variable_in_CASE_NOTE("  - AREP valid start date: " & arep_signature_date)
            call write_variable_in_CASE_NOTE("  - Client and AREP signed AREP form.")
            IF AREP_ID_check = checked THEN write_variable_in_CASE_NOTE("  - AREP ID on file.")
            IF TIKL_check = checked THEN write_variable_in_CASE_NOTE("  - TIKL'd for 12 months to get new HC AREP form.")
            If update_AREP_panel_checkbox = checked Then write_variable_in_CASE_NOTE("  - AREP panel updated.")
        End If
        call write_bullet_and_variable_in_case_note("ADDR", ADDR)
        call write_bullet_and_variable_in_case_note("FACI", FACI)
        call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
        call write_bullet_and_variable_in_case_note("DISA", DISA)
        If mof_form_checkbox = checked Then
            CALL write_variable_in_CASE_NOTE("* Medical Opinion Form Rec'd " & date_recd & " for M" & mof_hh_memb)
            IF mof_clt_release_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("  * Client signed release on MOF.")
            If last_exam_date <> "" Then CALL write_variable_in_CASE_NOTE("  - Date of last examination: " & last_exam_date)
            If doctor_date <> "" Then CALL write_variable_in_CASE_NOTE("  - Doctor signed form: " & doctor_date)
            If mof_time_condition_will_last <> "" Then  CALL write_variable_in_CASE_NOTE("  - Condition will last: " & mof_time_condition_will_last)
            If ability_to_work <> "" Then  CALL write_variable_in_CASE_NOTE("  - Ability to work: " & ability_to_work)
            If mof_other_notes <> "" Then  CALL write_variable_in_CASE_NOTE("  - Other notes: " & mof_other_notes)

            If SSA_application_indicated_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * The MOF indicates the client needs to apply for SSA.")
            If TTL_to_update_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * Specialized TTL team will review MOF and update the DISA panel as needed.")
            If TTL_email_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * An email regarding this MOF was sent to the TTL/FSSDataTeam for review on " & TTL_email_date & " by " & worker_signature & ".")
        End If
        call write_bullet_and_variable_in_case_note("JOBS", JOBS)
        If evf_form_received_checkbox = checked Then
            call write_variable_in_CASE_NOTE("* EVF received " & evf_date_recvd & ": " & EVF_status_dropdown & "*")
            Call write_variable_in_CASE_NOTE("  - Employer Name: " & employer)
            Call write_variable_in_CASE_NOTE("  - EVF for HH member: " & evf_ref_numb)
            'for additional information needed
            IF info = "yes" then
                Call write_variable_in_CASE_NOTE("  - Additional Info requested: " & info & " on " & info_date & " by " & request_info)
            	If EVF_TIKL_checkbox = checked then call write_variable_in_CASE_NOTE("* TIKL'd for 10 day return.")
            Else
                Call write_variable_in_CASE_NOTE("  - No additional information is needed/requested.")
            END IF
        End If
        call write_bullet_and_variable_in_case_note("BUSI", BUSI)
        call write_bullet_and_variable_in_case_note("BUSI/RBIC", BUSI_RBIC)
        call write_bullet_and_variable_in_case_note("UNEA", UNEA)
        If asset_form_checkbox = checked Then
            If LTC_case = vbNo Then
                Call write_variable_in_CASE_NOTE("* Signed Personal Statement about Assets for Cash Received (DHS 6054)")
                Call write_variable_in_CASE_NOTE("  - Received on: " & asset_form_doc_date)
                If signed_by_one <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_one & " on: " & signed_one_date)
                If signed_by_two <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_two & " on: " & signed_two_date)
                If signed_by_three <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_three & " on: " & signed_three_date)
                If box_one_info <> "" Then Call write_variable_in_CASE_NOTE("  - Account detail from form: " & box_one_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                        Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
                    End If
                Next
                If box_two_info <> "" Then Call write_variable_in_CASE_NOTE("  - Securities detail from form: " & box_two_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " Value: $" & ASSETS_ARRAY(ast_csv, the_asset) & " - Verif: " & left(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
                    End If
                Next
                If box_three_info <> "" Then Call write_variable_in_CASE_NOTE("  - Vehicle detail from form: " & box_three_info)
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & ASSETS_ARRAY(ast_owe_verif, the_asset))
                    End If
                Next
            End If

            If LTC_case = vbYes Then
                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                        Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
                    End If
                Next

                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                        If left(ASSETS_ARRAY(ast_type, the_asset), 2) <> "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset))
                        If left(ASSETS_ARRAY(ast_type, the_asset), 2) = "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset) & " LI Face Value: $" & ASSETS_ARRAY(ast_face_value, the_asset))
                        Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                    End If
                Next

                For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                    If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                        Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset))
                        Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                        If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_owe_verif, the_asset), len(ASSETS_ARRAY(ast_owe_verif, the_asset)) - 4))
                        If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                    End If
                Next
            End If
        End If
        call write_bullet_and_variable_in_case_note("ACCT", ACCT)
        call write_bullet_and_variable_in_case_note("SECU", SECU)
        call write_bullet_and_variable_in_case_note("CARS", CARS)
        call write_bullet_and_variable_in_case_note("REST", REST)
        call write_bullet_and_variable_in_case_note("Burial/OTHR", OTHR)
        call write_bullet_and_variable_in_case_note("Other assets", other_assets)
        call write_bullet_and_variable_in_case_note("SHEL", SHEL)
        call write_bullet_and_variable_in_case_note("INSA", INSA)
        call write_bullet_and_variable_in_case_note("Medical expenses", medical_expenses)
        call write_bullet_and_variable_in_case_note("Veteran's info", veterans_info)
        call write_bullet_and_variable_in_case_note("Other verifications", other_verifs)
        Call write_variable_in_case_note("---")
        call write_bullet_and_variable_in_case_note("Notes on your doc's", notes)
        call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
        IF HSR_scanner_checkbox = checked then Call write_variable_in_case_note("* Documents imaged to ECF.")
        call write_bullet_and_variable_in_case_note("Verifications still needed", verifs_needed)
    End If
    CALL write_variable_in_CASE_NOTE ("---")
    CALL write_variable_in_CASE_NOTE (worker_signature)
End If
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
If ltc_1503_form_checkbox = checked Then
    faci_footer_month = MAXIS_footer_month
    faci_footer_year = MAXIS_footer_year
    BeginDialog Dialog1, 0, 0, 365, 305, "1503 Dialog"
      EditBox 55, 5, 135, 15, FACI_1503
      DropListBox 255, 5, 95, 15, "30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", length_of_stay
      DropListBox 105, 25, 45, 15, "SNF"+chr(9)+"NF"+chr(9)+"ICF-DD"+chr(9)+"RTC", level_of_care
      DropListBox 215, 25, 135, 15, "acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-DD", admitted_from
      EditBox 145, 45, 205, 15, hospital_admitted_from
      EditBox 75, 65, 65, 15, admit_date
      EditBox 275, 65, 75, 15, discharge_date
      CheckBox 15, 85, 155, 10, "If you've processed this 1503, check here.", processed_1503_checkbox
      CheckBox 15, 115, 60, 10, "Updated RLVA?", updated_RLVA_checkbox
      CheckBox 85, 115, 60, 10, "Updated FACI?", updated_FACI_checkbox
      CheckBox 150, 115, 50, 10, "Need 3543?", need_3543_checkbox
      CheckBox 205, 115, 55, 10, "Need 3531?", need_3531_checkbox
      CheckBox 265, 115, 95, 10, "Need asset assessment?", need_asset_assessment_checkbox
      EditBox 130, 130, 225, 15, verifs_needed
      CheckBox 15, 150, 85, 10, "Sent 3050 back to LTCF", sent_3050_checkbox
      CheckBox 165, 155, 100, 10, "Sent verif req? If so, to who:", sent_verif_request_checkbox
      ComboBox 270, 150, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Client & AREP", sent_request_to
      CheckBox 15, 165, 120, 10, "Sent DHS-5181 to Case Manager", sent_5181_checkbox
      EditBox 30, 185, 330, 15, notes
      CheckBox 15, 215, 255, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", TIKL_checkbox
      CheckBox 15, 230, 155, 10, "Check here to have the script update HCMI.", HCMI_update_checkbox
      CheckBox 15, 245, 150, 10, "Check here to have the script update FACI.", FACI_update_checkbox
      EditBox 105, 265, 25, 15, faci_footer_month
      EditBox 135, 265, 25, 15, faci_footer_year
      EditBox 85, 285, 75, 15, mets_case_number
      ButtonGroup ButtonPressed
        OkButton 255, 285, 50, 15
        CancelButton 310, 285, 50, 15
      Text 5, 10, 50, 10, "Facility name:"
      Text 200, 10, 50, 10, "Length of stay:"
      Text 5, 30, 95, 10, "Recommended level of care:"
      Text 160, 30, 50, 10, "Admitted from:"
      Text 5, 50, 135, 10, "If hospital, list name/dates of admission:"
      Text 5, 70, 65, 10, "Date of admission:"
      Text 165, 70, 105, 10, "Date of discharge (if applicible):"
      GroupBox 0, 100, 360, 80, "Actions/Proofs"
      Text 10, 135, 115, 10, "Other proofs needed (if applicable):"
      Text 5, 190, 25, 10, "Notes:"
      GroupBox 5, 205, 355, 55, "Script actions"
      Text 5, 270, 95, 10, "Facility Update Month/Year:"
      Text 5, 290, 75, 10, "METS Case Number:"
    EndDialog

    Do
    	Do
            err_msg = ""
    		dialog Dialog1  					'Calling a dialog without a assigned variable will call the most recently defined dialog
    		Call cancel_continue_confirmation(skip_1503)
            If ButtonPressed = 0 then err_msg = "LOOP" & err_msg
            If skip_1503= TRUE Then
                err_msg = ""
                ltc_1503_form_checkbox = unchecked
            End If
    	LOOP UNTIL err_msg = ""        'currently there are no elements to review or mandate
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End If

If ltc_1503_form_checkbox = checked Then
    end_msg = end_msg & vbNewLine & "LTC 1503 Form information entered."
    Original_footer_month = MAXIS_footer_month
    Original_footer_year = MAXIS_footer_year
    MAXIS_footer_month = faci_footer_month
    MAXIS_footer_year = faci_footer_year
    'LTC 1503 gets it's own case note
    'navigating the script to the correct footer month
    back_to_self
    EMWriteScreen MAXIS_footer_month, 20, 43
    EMWriteScreen MAXIS_footer_year, 20, 46
    call navigate_to_MAXIS_screen("STAT", "FACI")

    'UPDATING MAXIS PANELS----------------------------------------------------------------------------------------------------
    'FACI
    If FACI_update_checkbox = checked then
    	call navigate_to_MAXIS_screen("stat", "faci")
    	EMReadScreen panel_max_check, 1, 2, 78
    	IF panel_max_check = "5" THEN
            stop_or_continue = MsgBox("This case has reached the maxzimum amount of FACI panels. Please review the case and delete an appropriate FACI panel." & vbNewLine & vbNewLine & "To continue the script run without updating FACI, press 'OK'." & vbNewLine & vbNewLine & "Otherwise, press 'CANCEL' to stop the script, and then rerunit with fewer than 5 FACI panels.", vbQuestion + vbOkCancel, "Continue without updating FACI?")
            If stop_or_continue = vbCancel Then script_end_procedure("~PT User Pressed Cancel")
            If stop_or_continue = vbOk Then FACI_update_checkbox = unchecked
    	ELSE
    		EMWriteScreen "nn", 20, 79
    		transmit
    	END IF
    End If
    If FACI_update_checkbox = checked then
        updated_FACI_checkbox = checked
    	EMWriteScreen FACI_1503, 6, 43
    	If level_of_care = "NF" then EMWriteScreen "42", 7, 43
    	If level_of_care = "RTC" THEN EMWriteScreen "47", 7, 43
    	If length_of_stay = "30 days or less" and level_of_care = "SNF" then EMWriteScreen "44", 7, 43
    	If length_of_stay = "31 to 90 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
    	If length_of_stay = "91 to 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
    	if length_of_stay = "over 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
    	If length_of_stay = "30 days or less" and level_of_care = "ICF-DD" then EMWriteScreen "44", 7, 43
    	If length_of_stay = "31 to 90 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
    	If length_of_stay = "91 to 180 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
    	If length_of_stay = "over 180 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
    	EMWriteScreen "n", 8, 43
    	Call create_MAXIS_friendly_date_with_YYYY(admit_date, 0, 14, 47)
    	If discharge_date<> "" then
    		Call create_MAXIS_friendly_date_with_YYYY(discharge_date, 0, 14, 71)
    		transmit
    		transmit
    	End if
    End if

    'HCMI
    If HCMI_update_checkbox = checked THEN
    	call navigate_to_MAXIS_screen("stat", "hcmi")
    	EMReadScreen HCMI_panel_check, 1, 2, 78
    	IF HCMI_panel_check <> "0" Then
    		PF9
    	ELSE
    		EMWriteScreen "nn", 20, 79
    		transmit
    	END IF
    	EMWriteScreen "dp", 10, 57
    	transmit
    	transmit
    END IF

    'THE TIKL----------------------------------------------------------------------------------------------------
    If length_of_stay = "30 days or less"   then TIKL_multiplier = 30
    If length_of_stay = "31 to 90 days"     then TIKL_multiplier = 90
    If length_of_stay = "91 to 180 days"    then TIKL_multiplier = 180
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    If TIKL_checkbox = checked then Call create_TIKL("Have " & worker_signature & " call " & FACI & " re: length of stay. " & TIKL_multiplier & " days expired.", TIKL_multiplier, admit_date, False, TIKL_note_text)

    'The CASE NOTE----------------------------------------------------------------------------------------------------
    Call start_a_blank_CASE_NOTE
    If processed_1503_checkbox = checked then
      	call write_variable_in_CASE_NOTE("***Processed 1503 from " & FACI_1503 & "***")
    Else
      	call write_variable_in_CASE_NOTE("***Rec'd 1503 from " & FACI_1503 & ", DID NOT PROCESS***")
    End if
    Call write_bullet_and_variable_in_case_note("Length of stay", length_of_stay)
    Call write_bullet_and_variable_in_case_note("Recommended level of care", level_of_care)
    Call write_bullet_and_variable_in_case_note("Admitted from", admitted_from)
    Call write_bullet_and_variable_in_case_note("Hospital admitted from", hospital_admitted_from)
    Call write_bullet_and_variable_in_case_note("Admit date", admit_date)
    Call write_bullet_and_variable_in_case_note("Discharge date", discharge_date)
    Call write_variable_in_CASE_NOTE("---")
    If updated_RLVA_checkbox = checked and updated_FACI_checkbox = checked then
    	Call write_variable_in_CASE_NOTE("* Updated RLVA and FACI.")
    Else
      	If updated_RLVA_checkbox = checked then Call write_variable_in_case_note("* Updated RLVA.")
      	If updated_FACI_checkbox = checked then Call write_variable_in_case_note("* Updated FACI.")
    End if
    If need_3543_checkbox = checked then Call write_variable_in_case_note("* A 3543 is needed for the client.")
    If need_3531_checkbox = checked then call write_variable_in_CASE_NOTE("* A 3531 is needed for the client.")
    If need_asset_assessment_checkbox = checked then call write_variable_in_CASE_NOTE("* An asset assessment is needed before a MA-LTC determination can be made.")
    If sent_3050_checkbox = checked then call write_variable_in_CASE_NOTE("* Sent 3050 back to LTCF.")
    If sent_5181_checkbox = checked then call write_variable_in_CASE_NOTE("* Sent DHS-5181 to Case Manager.")
    Call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
    If sent_verif_request_checkbox = checked then Call write_variable_in_case_note("* Sent verif request to " & sent_request_to)
    If processed_1503_checkbox = checked then Call write_variable_in_case_note("* Completed & Returned 1503 to LTCF.")
    If TIKL_checkbox = checked then Call write_variable_in_case_note("TIKL'd for " & TIKL_multiplier & " days to check length of stay.")
    Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", mets_case_number)
    Call write_bullet_and_variable_in_case_note("Notes", notes)
    Call write_variable_in_case_note("---")
    Call write_variable_in_case_note(worker_signature)
    MAXIS_footer_month = Original_footer_month
    MAXIS_footer_year = Original_footer_year
End If

If left(docs_rec, 2) = ", " Then docs_rec = right(docs_rec, len(docs_rec)-2)        'trimming the ',' off of the list of docs

If need_final_note = FALSE Then
    If ltc_1503_form_checkbox = checked Then script_end_procedure_with_error_report("The script run is complete, a case note for the LTC 1503 form has been entered. There are no additional documents indicated in the initial dialog and so the final note will not be entered as it would be blank.")
    If mtaf_form_checkbox = checked Then script_end_procedure_with_error_report("The script run is complete, a case note for the MTAF has been entered. The documents are either noted with the MTAF or there are no additional documents to be noted and the note would be blank.")
    script_end_procedure_with_error_report("The script run is complete, but no detail about documents has been added and the final case note will not be entered as it would be blank.")
End If

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
If actions_taken = "" Then
    BeginDialog Dialog1, 0, 0, 251, 70, "Actions Taken Dialog"
      ButtonGroup ButtonPressed
        OkButton 195, 50, 50, 15
      Text 5, 10, 205, 10, "The actions taken have not been detailed. Explain them here:"
      EditBox 10, 25, 235, 15, actions_taken
    EndDialog

    Do
        Do
            err_msg = ""
            dialog Dialog1
            cancel_confirmation
            actions_taken = trim(actions_taken)
            If actions_taken = "" Then err_msg = err_msg & vbNewLine & "* Enter the actions taken."
            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE
End If

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).

call start_a_blank_CASE_NOTE
If HSR_scanner_checkbox = checked then
    Call write_variable_in_case_note("Docs Rec'd & scanned: " & docs_rec)
else
    Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
END IF
call write_bullet_and_variable_in_case_note("Document date stamp", doc_date_stamp)
If arep_form_checkbox = checked Then
    call write_variable_in_CASE_NOTE("* AREP FORM received on " & AREP_recvd_date & ". AREP: " & arep_name)
    If dhs_3437_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the DHS 3437 - MHCP AUTHORIZED REPRESENTATIVE REQUEST Form.")
    If HC_12729_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named on the HC 12729 - AUTHORIZED REPRESENTATIVE REQUEST Form.")
    If D405_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("  - AREP name on the SNAP AUTHORIZED REPRESENTATIVE CHOICE D405 Form.")
        Call write_variable_in_CASE_NOTE("  - AREP also authorixed to get and use EBT Card.")
    End If
    If CAF_AREP_page_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in the CAF.")
    If HCAPP_AREP_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP named in a Health Care Application.")
    If power_of_attorney_checkbox = checked Then Call write_variable_in_CASE_NOTE("  - AREP has Power of Attorney Designation.")
    If AREP_programs <> "" then call write_variable_in_CASE_NOTE("  - Programs Authorized for: " & AREP_programs)
    If arep_signature_date <> "" Then call write_variable_in_CASE_NOTE("  - AREP valid start date: " & arep_signature_date)
    call write_variable_in_CASE_NOTE("  - Client and AREP signed AREP form.")
    IF AREP_ID_check = checked THEN write_variable_in_CASE_NOTE("  - AREP ID on file.")
    IF TIKL_check = checked THEN write_variable_in_CASE_NOTE("  - TIKL'd for 12 months to get new HC AREP form.")
    If update_AREP_panel_checkbox = checked Then write_variable_in_CASE_NOTE("  - AREP panel updated.")
End If
call write_bullet_and_variable_in_case_note("ADDR", ADDR)
call write_bullet_and_variable_in_case_note("FACI", FACI)
call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
call write_bullet_and_variable_in_case_note("DISA", DISA)
If mof_form_checkbox = checked Then
    CALL write_variable_in_CASE_NOTE("* Medical Opinion Form Rec'd " & date_recd & " for M" & mof_hh_memb)
    IF mof_clt_release_checkbox = checked THEN CALL write_variable_in_CASE_NOTE ("  - *Client signed release on MOF.*")
    CALL write_variable_in_CASE_NOTE("  - Date of last examination: " & last_exam_date)
    CALL write_variable_in_CASE_NOTE("  - Doctor signed form: " & doctor_date)
    CALL write_variable_in_CASE_NOTE("  - Condition will last: " & mof_time_condition_will_last)
    CALL write_variable_in_CASE_NOTE("  - Ability to work: " & ability_to_work)
    CALL write_variable_in_CASE_NOTE("  - Other notes: " & mof_other_notes)
    If SSA_application_indicated_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * The MOF indicates the client needs to apply for SSA.")
    If TTL_to_update_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * Specialized TTL team will review MOF and update the DISA panel as needed.")
    If TTL_email_checkbox = checked Then Call write_variable_in_CASE_NOTE("  * An email regarding this MOF was sent to the TTL/FSSDataTeam for review on " & TTL_email_date & " by " & worker_signature & ".")
End If
call write_bullet_and_variable_in_case_note("JOBS", JOBS)
If evf_form_received_checkbox = checked Then
    call write_variable_in_CASE_NOTE("* EVF received " & evf_date_recvd & ": " & EVF_status_dropdown & "*")
    Call write_variable_in_CASE_NOTE("  - Employer Name: " & employer)
    Call write_variable_in_CASE_NOTE("  - EVF for HH member: " & evf_ref_numb)
    'for additional information needed
    IF info = "yes" then
        Call write_variable_in_CASE_NOTE("  - Additional Info requested: " & info & " on " & info_date & " by " & request_info)
    	If EVF_TIKL_checkbox = checked then call write_variable_in_CASE_NOTE("* TIKL'd for 10 day return.")
    Else
        Call write_variable_in_CASE_NOTE("  - No additional information is needed/requested.")
    END IF
End If
call write_bullet_and_variable_in_case_note("BUSI", BUSI)
call write_bullet_and_variable_in_case_note("BUSI/RBIC", BUSI_RBIC)
call write_bullet_and_variable_in_case_note("UNEA", UNEA)
If asset_form_checkbox = checked Then
    If LTC_case = vbNo Then
        Call write_variable_in_CASE_NOTE("* Signed Personal Statement about Assets for Cash Received (DHS 6054)")
        Call write_variable_in_CASE_NOTE("  - Received on: " & asset_form_doc_date)
        If signed_by_one <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_one & " on: " & signed_one_date)
        If signed_by_two <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_two & " on: " & signed_two_date)
        If signed_by_three <> "Select or Type" Then Call write_variable_in_CASE_NOTE("  - Signed by: " & signed_by_three & " on: " & signed_three_date)
        If box_one_info <> "" Then Call write_variable_in_CASE_NOTE("  - Account detail from form: " & box_one_info)
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
            End If
        Next
        If box_two_info <> "" Then Call write_variable_in_CASE_NOTE("  - Securities detail from form: " & box_two_info)
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " Value: $" & ASSETS_ARRAY(ast_csv, the_asset) & " - Verif: " & left(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
            End If
        Next
        If box_three_info <> "" Then Call write_variable_in_CASE_NOTE("  - Vehicle detail from form: " & box_three_info)
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & ASSETS_ARRAY(ast_owe_verif, the_asset))
            End If
        Next
    End If

    If LTC_case = vbYes Then
        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. At: " & ASSETS_ARRAY(ast_location, the_asset))
                Call write_variable_in_CASE_NOTE("      Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
                If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("      " & ASSETS_ARRAY(ast_share_note, the_asset))
            End If
        Next

        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                If left(ASSETS_ARRAY(ast_type, the_asset), 2) <> "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset))
                If left(ASSETS_ARRAY(ast_type, the_asset), 2) = "LI" Then Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " CSV: $" & ASSETS_ARRAY(ast_csv, the_asset) & " LI Face Value: $" & ASSETS_ARRAY(ast_face_value, the_asset))
                Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Security is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the security.")
                If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
            End If
        Next

        For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
            If ASSETS_ARRAY(cnote_panel, the_asset) = checked AND  ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & ": " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " - Trade-In Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset))
                Call write_variable_in_CASE_NOTE("      Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4) & " - Rec'vd On: " & ASSETS_ARRAY(ast_verif_date, the_asset))
                If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_owe_verif, the_asset), len(ASSETS_ARRAY(ast_owe_verif, the_asset)) - 4))
                If ASSETS_ARRAY(ast_note, the_asset) <> "" Then Call write_variable_in_CASE_NOTE("      Notes: " & ASSETS_ARRAY(ast_note, the_asset))
            End If
        Next
    End If
End If
call write_bullet_and_variable_in_case_note("ACCT", ACCT)
call write_bullet_and_variable_in_case_note("SECU", SECU)
call write_bullet_and_variable_in_case_note("CARS", CARS)
call write_bullet_and_variable_in_case_note("REST", REST)
call write_bullet_and_variable_in_case_note("Burial/OTHR", OTHR)
call write_bullet_and_variable_in_case_note("Other assets", other_assets)
call write_bullet_and_variable_in_case_note("SHEL", SHEL)
call write_bullet_and_variable_in_case_note("INSA", INSA)
call write_bullet_and_variable_in_case_note("Medical expenses", medical_expenses)
call write_bullet_and_variable_in_case_note("Veteran's info", veterans_info)
call write_bullet_and_variable_in_case_note("Other verifications", other_verifs)
Call write_variable_in_case_note("---")
call write_bullet_and_variable_in_case_note("Notes on your doc's", notes)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
IF HSR_scanner_checkbox = checked then Call write_variable_in_case_note("* Documents imaged to ECF.")
call write_bullet_and_variable_in_case_note("Verifications still needed", verifs_needed)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

Call confirm_docs_accepted_in_ecf(end_msg)      'function that asks if ECF documents have been accepted

script_end_procedure_with_error_report(end_msg)
