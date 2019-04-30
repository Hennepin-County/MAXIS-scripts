'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DOCUMENTS RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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

'This function creates the HH Member dropdown for a number of different dialogs
function Generate_Client_List(list_for_dropdown)

	memb_row = 5       'setting the row to look at the list of members on the left hand side of the panel

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")         'go to MEMB
	Do                                                     'this loop transmits to each MEMB panel to read information for each member
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do           'this is the end of the list of members
		EMWriteScreen ref_numb, 20, 76            'writing the reference number in the command line to go to each MEMB panel
		transmit
		EMReadScreen first_name, 12, 6, 63        'reading the name on the panel
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")     'adding each client information to a string
		memb_row = memb_row + 1                   'going to the next member
	Loop until memb_row = 20

    If memb_row = 6 Then        'If the row is only 6, then there is only one person in the HH
        list_for_dropdown = right(client_info, len(client_info) - 1)    'taking the '~' off of the string
    Else
    	client_info = right(client_info, len(client_info) - 1)             'taking the left most '~' off
    	client_list_array = split(client_info, "~")                        'making this an array

    	For each person in client_list_array                               'creating the string to be added to the dialog code to fill the dropdown
    		list_for_dropdown = list_for_dropdown & chr(9) & person
    	Next
    End If

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

Const update_panel      = 42

Dim client_list_array

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
'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel or VbDefaultButton2) 'defaults to no since that is most commonly chosen option
If LTC_case = vbCancel then stopscript

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

Call get_county_code()
end_msg = ""

'Displays the dialog and navigates to case note
'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
DO
	Do
        err_msg = ""

		If LTC_case = vbNo then
            BeginDialog Dialog1, 0, 0, 416, 395, "Documents received"       'This is the regular (NON LTC) dialog
              EditBox 80, 5, 60, 15, MAXIS_case_number
              EditBox 225, 5, 60, 15, doc_date_stamp
              If worker_county_code = "x127" Then  CheckBox 355, 10, 60, 10, "OTS scanning", HSR_scanner_checkbox
              EditBox 80, 25, 330, 15, docs_rec
              EditBox 35, 70, 315, 15, ADDR
              EditBox 75, 90, 275, 15, SCHL
              EditBox 35, 110, 315, 15, DISA
              EditBox 35, 130, 315, 15, JOBS
              CheckBox 370, 115, 30, 10, "MOF", mof_form_checkbox
              CheckBox 370, 135, 30, 10, "EVF", evf_form_received_checkbox
              CheckBox 370, 195, 30, 10, "Asset", asset_form_checkbox
              Text 370, 205, 35, 10, "Statement"
              CheckBox 370, 220, 30, 10, "AREP", arep_form_checkbox
              'CheckBox 365, 260, 40, 10, "LTC1503", ltc_1503_form_checkbox
              EditBox 35, 150, 315, 15, BUSI
              EditBox 35, 170, 315, 15, UNEA
              EditBox 35, 190, 315, 15, ACCT
              EditBox 60, 210, 290, 15, other_assets
              EditBox 35, 230, 315, 15, SHEL
              EditBox 35, 250, 315, 15, INSA
              EditBox 55, 270, 295, 15, other_verifs
              EditBox 80, 310, 320, 15, notes
              EditBox 80, 330, 320, 15, actions_taken
              EditBox 80, 350, 320, 15, verifs_needed
              EditBox 220, 375, 80, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 305, 375, 50, 15
                CancelButton 360, 375, 50, 15
              Text 10, 115, 25, 10, "DISA:"
              Text 10, 135, 25, 10, "JOBS:"
              Text 10, 155, 20, 10, "BUSI:"
              Text 10, 175, 25, 10, "UNEA:"
              Text 10, 195, 25, 10, "ACCT:"
              Text 10, 215, 45, 10, "Other assets:"
              Text 10, 235, 25, 10, "SHEL:"
              Text 10, 255, 20, 10, "INSA:"
              Text 10, 275, 45, 10, "Other verif's:"
              Text 155, 380, 60, 10, "Worker signature:"
              Text 10, 75, 25, 10, "ADDR:"
              Text 10, 315, 70, 10, "Notes on your doc's:"
              Text 30, 10, 45, 10, "Case number:"
              Text 10, 335, 50, 10, "Actions taken:"
              Text 140, 45, 205, 10, "Note: What you enter above will become the case note header."
              Text 10, 30, 70, 10, "Documents received: "
              Text 150, 10, 75, 10, "Document date stamp:"
              Text 10, 355, 65, 10, "Verif's still needed:"
              GroupBox 5, 55, 350, 235, "Breakdown of Documents received"
              GroupBox 5, 295, 405, 75, "Additional information"
              Text 10, 95, 65, 10, "SCHL/STIN/STEC:"
              GroupBox 360, 55, 50, 235, "FORMS"
              Text 370, 65, 35, 45, "Watch for more form options - coming soon!"
            EndDialog

        ElseIf LTC_case = vbYes then
            BeginDialog Dialog1, 0, 0, 416, 425, "Documents received LTC"           'This is the LTC Dialog
              EditBox 80, 5, 60, 15, MAXIS_case_number
              EditBox 230, 5, 60, 15, doc_date_stamp
              If worker_county_code = "x127" Then  CheckBox 355, 10, 60, 10, "OTS scanning", HSR_scanner_checkbox
              EditBox 80, 25, 330, 15, docs_rec
              EditBox 35, 65, 315, 15, FACI
              EditBox 35, 85, 135, 15, JOBS
              EditBox 215, 85, 135, 15, BUSI_RBIC
              CheckBox 370, 90, 30, 10, "EVF", evf_form_received_checkbox
              CheckBox 370, 130, 30, 10, "Asset", asset_form_checkbox
              Text 370, 140, 35, 10, "Statement"
              'CheckBox 370, 115, 30, 10, "MOF", mof_form_checkbox
              CheckBox 370, 280, 30, 10, "AREP", arep_form_checkbox
              CheckBox 365, 295, 40, 10, "LTC1503", ltc_1503_form_checkbox
              EditBox 35, 105, 315, 15, UNEA
              EditBox 35, 125, 315, 15, ACCT
              EditBox 35, 145, 315, 15, SECU
              EditBox 35, 165, 315, 15, CARS
              EditBox 35, 185, 315, 15, REST
              EditBox 65, 205, 285, 15, OTHR
              EditBox 35, 225, 315, 15, SHEL
              EditBox 35, 245, 315, 15, INSA
              EditBox 80, 265, 270, 15, medical_expenses
              EditBox 55, 285, 295, 15, veterans_info
              EditBox 55, 305, 295, 15, other_verifs
              EditBox 80, 340, 330, 15, notes
              EditBox 80, 360, 330, 15, actions_taken
              EditBox 80, 380, 330, 15, verifs_needed
              EditBox 225, 405, 80, 15, worker_signature
              ButtonGroup ButtonPressed
                OkButton 310, 405, 50, 15
                CancelButton 360, 405, 50, 15
              Text 10, 170, 20, 10, "CARS:"
              Text 10, 190, 20, 10, "REST:"
              Text 10, 210, 50, 10, "BURIAL/OTHR:"
              Text 10, 230, 25, 10, "SHEL:"
              Text 10, 250, 25, 10, "INSA:"
              Text 10, 310, 45, 10, "Other verif's:"
              Text 165, 410, 60, 10, "Worker signature:"
              Text 10, 70, 25, 10, "FACI:"
              Text 10, 345, 70, 10, "Notes on your doc's:"
              Text 30, 10, 50, 10, "Case number:"
              Text 10, 365, 50, 10, "Actions taken:"
              Text 205, 40, 205, 10, "Note: What you enter above will become the case note header."
              Text 5, 30, 70, 10, "Documents received: "
              Text 155, 10, 75, 10, "Document date stamp:"
              Text 10, 385, 70, 10, "Verif's still needed:"
              GroupBox 5, 50, 350, 275, "Breakdown of Documents received"
              Text 10, 130, 20, 10, "ACCT:"
              Text 175, 90, 40, 10, "BUSI/RBIC:"
              Text 10, 110, 25, 10, "UNEA:"
              Text 10, 290, 45, 10, "Veteran info:"
              Text 10, 90, 20, 10, "JOBS:"
              Text 10, 270, 65, 10, "Medical expenses:"
              GroupBox 5, 330, 410, 70, "Additional information"
              Text 10, 150, 20, 10, "SECU:"
              GroupBox 360, 55, 50, 275, "FORMS"
              Text 370, 165, 35, 45, "Watch for more form options - coming soon!"
            EndDialog
        End If

        dialog Dialog1
		cancel_confirmation																'quits if cancel is pressed

        Call validate_MAXIS_case_number(err_msg, "*")
		If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* You must sign your case note."
        If HSR_scanner_checkbox = unchecked and actions_taken = "" Then
            If evf_form_received_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* You must case note your actions taken."
        End If

        If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

	LOOP until err_msg = ""													'Loops until that case number exists
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

CALL Generate_Client_List(client_dropdown)

If LTC_case = vbNo then end_msg = "Sucess! Documents received noted for case."
If LTC_case = vbYes then end_msg = "Sucess! Documents received noted for LTC case."

'This will be for any functionality that needs the HH Member array
If asset_form_checkbox = checked Then
    call HH_member_custom_dialog(HH_member_array)
End If

'EVF HANDLING =======================================================================================
If evf_form_received_checkbox = checked Then
    EVF_TIKL_checkbox = checked 'defaulting the TIKL checkbox to be checked initially in the dialog.
    evf_date_recvd = doc_date_stamp

    BeginDialog Dialog1, 0, 0, 291, 205, "Employment Verification Form Received"
      Text 70, 10, 60, 10, MAXIS_case_number
      EditBox 220, 5, 60, 15, evf_date_recvd
      ComboBox 70, 30, 210, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
      EditBox 70, 50, 210, 15, employer
      DropListBox 70, 75, 210, 45, "Select One..."+chr(9)+client_dropdown, evf_client
      DropListBox 75, 110, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", info
      EditBox 220, 110, 60, 15, info_date
      EditBox 75, 130, 60, 15, request_info
      CheckBox 160, 135, 105, 10, "10 day TIKL for additional info", EVF_TIKL_checkbox
      EditBox 70, 160, 210, 15, actions_taken
      EditBox 70, 180, 100, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 175, 180, 50, 15
        CancelButton 230, 180, 50, 15
      Text 10, 135, 65, 10, "Info Requested via:"
      Text 10, 115, 60, 10, "Addt'l Info Reqstd:"
      Text 5, 75, 60, 10, "Household Memb:"
      Text 10, 55, 55, 10, "Employer name:"
      Text 15, 165, 50, 10, "Actions taken:"
      Text 5, 185, 60, 10, "Worker Signature:"
      Text 25, 35, 40, 10, "EVF Status:"
      Text 150, 10, 65, 10, "Date EVF received:"
      Text 15, 10, 50, 10, "Case Number:"
      Text 160, 115, 55, 10, "Date Requested:"
      GroupBox 5, 95, 280, 60, "Is additional information needed?"
    EndDialog

    'starts the EVF received case note dialog
    DO
    	Do
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
    		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered
    		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!" 		'checks that the case note was signed
            If skip_evf = TRUE Then
                evf_form_received_checkbox = unchecked
                err_msg = ""
                EVF_TIKL_checkbox = unchecked
            End If
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "* Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
End If

if evf_form_received_checkbox = checked Then
    evf_ref_numb = left(evf_client, 2)
    docs_rec = docs_rec & ", EVF for M" & evf_ref_numb

    'Checks if additional info is yes and the TIKL is checked, sets a TIKL for the return of the info
    If EVF_TIKL_checkbox = checked Then
    	call navigate_to_MAXIS_screen("dail", "writ")
    	call create_MAXIS_friendly_date(date, 10, 5, 18)		'The following will generate a TIKL formatted date for 10 days from now.
    	call write_variable_in_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action. (TIKL auto-generated from script)." )
    	transmit
    	PF3
    	'Success message
    	end_msg = end_msg & vbNewLine & "Additional detail added about EVF." & vbNewLine & "TIKL has been sent for 10 days from now for the additional information requested."
    Else
        end_msg = end_msg & vbNewLine & "Additional detail added about EVF."
    End If
End If

If mof_form_checkbox = checked Then
    mof_date_recd = doc_date_stamp

    BeginDialog Dialog1, 0, 0, 216, 185, "Medical Opinion Form Received for Case #" & MAXIS_case_number
      EditBox 55, 5, 50, 15, mof_date_recd
      CheckBox 125, 10, 85, 10, "Client signed release?", mof_clt_release_checkbox
      DropListBox 80, 25, 130, 45, "Select One..."+chr(9)+client_dropdown, mof_hh_memb
      EditBox 90, 45, 55, 15, last_exam_date
      EditBox 90, 65, 55, 15, doctor_date
      ComboBox 70, 85, 140, 45, " "+chr(9)+"Less than 30 Days"+chr(9)+"Between 30 - 45 Days"+chr(9)+"More than 45 Days"+chr(9)+"No End Date Listed", mof_time_condition_will_last
      EditBox 85, 105, 125, 15, ability_to_work
      EditBox 55, 125, 155, 15, mof_other_notes
      EditBox 55, 145, 155, 15, actions_taken
      ButtonGroup ButtonPressed
        OkButton 105, 165, 50, 15
        CancelButton 160, 165, 50, 15
      Text 5, 10, 50, 10, "Date received: "
      Text 5, 30, 70, 10, "HHLD Member name"
      Text 5, 50, 60, 10, "Date of last exam: "
      Text 5, 70, 80, 10, "Date doctor signed form: "
      Text 5, 90, 60, 10, "Condition will last:"
      Text 5, 110, 75, 10, "Client's ability to work: "
      Text 5, 130, 40, 10, "Other notes: "
      Text 5, 150, 45, 10, "Action taken: "
      Text 155, 45, 50, 35, "Do not enter diagnosis in case notes per PQ #16506."
    EndDialog

    Do
        DO
        	Err_msg = ""
        	Dialog Dialog1
        	Call cancel_continue_confirmation(skip_MOF)
            'Call validate_MAXIS_case_number(err_msg, "*")
            If mof_hh_memb = "Select One..." Then err_msg = err_ms & vbNewLine & "* Select the household member."
            If skip_MOF= TRUE Then
                err_msg = ""
                mof_form_checkbox = unchecked
            End If
        	If err_msg <> "" Then msgbox "Please resolve the following for the script to continue:" & vbNewLine & err_msg
        LOOP until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

End If

If mof_form_checkbox = checked Then
    mof_ref_numb = left(mof_hh_memb, 2)
    docs_rec = docs_rec & ", MOF for M" & mof_ref_numb
End If


If asset_form_checkbox = checked Then
    asset_counter = 0
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
                If SECU_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Agency Form - 1"
                If SECU_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Source Doc - 2"
                If SECU_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Phone Contact - 3"
                If SECU_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Other Document - 5"
                If SECU_verif = "6" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Personal Stmt - 6"
                If SECU_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "No Ver Prov - N"
                ASSETS_ARRAY(ast_face_value, asset_counter) = trim(SECU_face_value)
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
                If CARS_type = "1" Then ASSETS_ARRAY(ast_type, asset_counter) = "Car - 1"
                If CARS_type = "2" Then ASSETS_ARRAY(ast_type, asset_counter) = "Truck - 2"
                If CARS_type = "3" Then ASSETS_ARRAY(ast_type, asset_counter) = "Van - 3"
                If CARS_type = "4" Then ASSETS_ARRAY(ast_type, asset_counter) = "Camper - 4"
                If CARS_type = "5" Then ASSETS_ARRAY(ast_type, asset_counter) = "Motorcycle - 5"
                If CARS_type = "6" Then ASSETS_ARRAY(ast_type, asset_counter) = "Trailer - 6"
                If CARS_type = "7" Then ASSETS_ARRAY(ast_type, asset_counter) = "Other - 7"
                ASSETS_ARRAY(ast_make, asset_counter) = replace(CARS_make, "_", "")
                ASSETS_ARRAY(ast_model, asset_counter) = replace(CARS_model, "_", "")
                ASSETS_ARRAY(ast_trd_in, asset_counter) = trim(CARS_trade_in)
                ASSETS_ARRAY(ast_loan_value, asset_counter) = trim(CARS_loan)
                If CARS_source = "1" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "NADA - 1"
                If CARS_source = "2" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "Appraisal Val - 2"
                If CARS_source = "3" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "Client Stmt - 3"
                If CARS_source = "4" Then ASSETS_ARRAY(ast_value_srce, asset_counter) = "Other Document - 4"
                If CARS_owner_verif = "1" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Title - 1"
                If CARS_owner_verif = "2" Then ASSETS_ARRAY(ast_verif, asset_counter) = "License Reg - 2"
                If CARS_owner_verif = "3" Then ASSETS_ARRAY(ast_verif, asset_counter) = "DMV - 3"
                If CARS_owner_verif = "4" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Purchase Agmt - 4"
                If CARS_owner_verif = "5" Then ASSETS_ARRAY(ast_verif, asset_counter) = "Other Document - 5"
                If CARS_owner_verif = "N" Then ASSETS_ARRAY(ast_verif, asset_counter) = "No Ver Prvd - N"
                ASSETS_ARRAY(ast_amt_owed, asset_counter) = trim(replace(CARS_owe_amount, "_", ""))
                ASSETS_ARRAY(ast_owe_YN, asset_counter) = replace(CARS_joint_owner_YN, "_", "")
                ASSETS_ARRAY(ast_bal_date, asset_counter) = replace(CARS_owed_date, " ", "/")
                If ASSETS_ARRAY(ast_bal_date, asset_counter) = "__/__/__" Then ASSETS_ARRAY(ast_bal_date, asset_counter) = ""
                If CARS_use = "1" Then ASSETS_ARRAY(ast_use, asset_counter) = "Primary Veh - 1"
                If CARS_use = "2" Then ASSETS_ARRAY(ast_use, asset_counter) = "Emp/Trng Trans/Emp Search - 2"
                If CARS_use = "3" Then ASSETS_ARRAY(ast_use, asset_counter) = "Disa Trans - 3"
                If CARS_use = "4" Then ASSETS_ARRAY(ast_use, asset_counter) = "Inc Producing - 4"
                If CARS_use = "5" Then ASSETS_ARRAY(ast_use, asset_counter) = "Used As Home - 5"
                If CARS_use = "7" Then ASSETS_ARRAY(ast_use, asset_counter) = "Unlicensed - 7"
                If CARS_use = "8" Then ASSETS_ARRAY(ast_use, asset_counter) = "Othr Countable - 8"
                If CARS_use = "9" Then ASSETS_ARRAY(ast_use, asset_counter) = "Unavailable - 9"
                If CARS_use = "0" Then ASSETS_ARRAY(ast_use, asset_counter) = "Long Distance Emp Travel - 0"
                If CARS_use = "A" Then ASSETS_ARRAY(ast_use, asset_counter) = "Carry Heating Fuel Or Water - A"
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

        MsgBox asset_display

        dlg_len = 240

        If acct_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * acct_panels)
        If secu_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * secu_panels)
        If cars_panels > 0 Then dlg_len = dlg_len + 15
        dlg_len = dlg_len + (10 * cars_panels)
        MsgBox dlg_len

        y_pos = 40
        BeginDialog Dialog1, 0, 0, 390, dlg_len, "Signed Personal Statement about Assest for Case #" & MAXIS_case_number
          Text 10, 10, 265, 10, "Assets for SNAP/Cash are self attested and are reported on this form (DHS 6054)."
          Text 10, 30, 95, 10, "Date the form was received:"
          EditBox 110, 25, 35, 15, asset_form_doc_date
          If acct_panels > 0 Then
              Text 10, y_pos, 95, 10, "Current ACCT panel details."
              Text 240, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "ACCT" Then
                      Text 15, y_pos, 210, 10,  "* ACCT " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset) & " - Balance: $" & ASSETS_ARRAY(ast_balance, the_asset)
                      CheckBox 240, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
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
              Text 240, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "SECU" Then
                      Text 15, y_pos, 210, 10, "* SECU " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_type, the_asset) & " @ " & ASSETS_ARRAY(ast_location, the_asset)
                      CheckBox 240, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
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
              Text 240, y_pos, 120, 10, "Check to include in CASE/NOTE"
              y_pos = y_pos + 15
              For the_asset = 0 to Ubound(ASSETS_ARRAY, 2)
                  If ASSETS_ARRAY(ast_panel, the_asset) = "CARS" Then
                      Text 15, y_pos, 210, 10, "* CARS " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & ASSETS_ARRAY(ast_instance, the_asset) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset)
                      CheckBox 240, y_pos, 45, 10, "Updated", ASSETS_ARRAY(cnote_panel, the_asset)
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
          ComboBox 10, y_pos, 105, 45, "Select or Type"+chr(9)+client_dropdown, signed_by_one
          EditBox 130, y_pos, 50, 15, signed_one_date
          y_pos = y_pos + 20
          ComboBox 10, y_pos, 105, 45, "Select or Type"+chr(9)+client_dropdown, signed_by_two
          EditBox 130, y_pos, 50, 15, signed_two_date
          y_pos = y_pos + 20
          ComboBox 10, y_pos, 105, 45, "Select or Type"+chr(9)+client_dropdown, signed_by_three
          EditBox 130, y_pos, 50, 15, signed_three_date

          CheckBox 240, y_pos_over, 130, 10, "Check here to have the script update", run_updater_checkbox
          Text 250, y_pos_over + 10, 115, 10, "asset panels. (ACCT, SECU, CARS)."
          Text 255, y_pos_over + 25, 85, 20, "*Panels update by the script will be case noted."
          ButtonGroup ButtonPressed
            OkButton 280, y_pos, 50, 15
            CancelButton 335, y_pos, 50, 15
        EndDialog

        Do
            Do
                err_msg = ""

                dialog Dialog1

                Call cancel_continue_confirmation(skip_asset)
                If skip_asset= TRUE Then
                    err_msg = ""
                    asset_form_checkbox = unchecked
                End If

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

            Loop Until err_mag = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE


    Else

    End If

End If

highest_asset_count = asset_counter

If asset_form_checkbox = checked Then
    If LTC_case = vbNo Then docs_rec = docs_rec & ", Personal Statement (DHS 6054)"
    If LTC_case = vbYes Then docs_rec = docs_rec & ", Asset documents"
    If run_updater_checkbox = checked Then
        MAXIS_footer_month = CM_mo
        MAXIS_footer_year = CM_yr

        Do
            found_the_panel = FALSE
            asset_counter = highest_asset_count
            update_panel_type = "NONE - I'm all done"
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

                Loop until err_msg = ""
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

            panel_type = right(update_panel_type, 4)
            skip_this_panel = FALSE

            If panel_type = "ACCT" Then
                If update_panel_type = "Existing ACCT" Then
                    Call navigate_to_MAXIS_screen("STAT", "ACCT")
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
                        MsgBox "Current member: " & current_member & vbNewLine & "Array member: " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & vbNewLine & "Current instance: " & current_instance & vbNewLine & "Array instance: " & ASSETS_ARRAY(ast_instance, the_asset)
                        If current_member = ASSETS_ARRAY(ast_ref_nbr, the_asset) AND current_instance = ASSETS_ARRAY(ast_instance, the_asset) Then
                            asset_counter = the_asset
                            If ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y" Then count_cash_checkbox = checked
                            If ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y" Then count_snap_checkbox = checked
                            If ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y" Then count_hc_checkbox = checked
                            If ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y" Then count_grh_checkbox = checked
                            If ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y" Then count_ive_checkbox = checked
                            MsgBox ASSETS_ARRAY(ast_own_ratio, asset_counter)
                            share_ratio_num = left(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            share_ratio_denom = right(ASSETS_ARRAY(ast_own_ratio, asset_counter), 1)
                            Exit For
                        End If
                    Next

                Else update_panel_type = "New ACCT"
                    ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                End If

                If share_ratio_num = "" Then share_ratio_num = "1"
                If share_ratio_denom = "" Then share_ratio_denom = "1"
                If LTC_case = vbNo AND ASSETS_ARRAY(ast_verif, asset_counter) = "" Then ASSETS_ARRAY(ast_verif, asset_counter) = "6 - Personal Statement"

                'Dialog to fill the ACCT panel
                BeginDialog Dialog1, 0, 0, 271, 235, "New ACCT panel for Case #" & MAXIS_case_number
                  DropListBox 75, 10, 135, 45, "Select One..."+chr(9)+client_dropdown, ASSETS_ARRAY(ast_owner, asset_counter)
                  DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+"SV - Savings"+chr(9)+"CK - Checking"+chr(9)+"CE - Certificate of Deposit"+chr(9)+"MM - Money Market"+chr(9)+"DC - Debit Card"+chr(9)+"KO - Keogh Account"+chr(9)+"FT - Federal Thrift Savings Plan"+chr(9)+"SL - State and Local Govt Ret"+chr(9)+"RA - Employee Ret Annuities"+chr(9)+"NP - Non-Profit Employer Ret Plans"+chr(9)+"IR - Indiv Ret Acct"+chr(9)+"RH - Roth IRA"+chr(9)+"FR - Ret Plans for Certain Employees"+chr(9)+"CT - Corp Ret Trust (before 1959)"+chr(9)+"RT - Other Ret Fund"+chr(9)+"QT - Qualified Tuition (529)"+chr(9)+"CA - Coverdell SV (530)"+chr(9)+"OE - Other Educationsal"+chr(9)+"OT - Other Account Type", ASSETS_ARRAY(ast_type, asset_counter)
                  EditBox 75, 50, 105, 15, ASSETS_ARRAY(ast_number, asset_counter)
                  EditBox 75, 70, 105, 15, ASSETS_ARRAY(ast_location, asset_counter)
                  EditBox 75, 90, 50, 15, ASSETS_ARRAY(ast_balance, asset_counter)
                  EditBox 160, 90, 50, 15, ASSETS_ARRAY(ast_bal_date, asset_counter)
                  DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_verif, asset_counter)
                  CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
                  CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
                  CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
                  CheckBox 230, 70, 25, 10, "GRH", count_grh_checkbox
                  CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
                  EditBox 75, 145, 50, 15, ASSETS_ARRAY(ast_wdrw_penlty, asset_counter)
                  DropListBox 75, 165, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ASSETS_ARRAY(ast_wthdr_verif, asset_counter)
                  EditBox 215, 125, 15, 15, share_ratio_num
                  EditBox 240, 125, 15, 15, share_ratio_denom
                  ComboBox 170, 160, 90, 45, "Type or Select"+chr(9)+client_dropdown, ASSETS_ARRAY(ast_othr_ownr_one, asset_counter)
                  ComboBox 170, 175, 90, 45, "Type or Select"+chr(9)+client_dropdown, ASSETS_ARRAY(ast_othr_ownr_two, asset_counter)
                  ComboBox 170, 190, 90, 45, "Type or Select"+chr(9)+client_dropdown, ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter)
                  EditBox 75, 190, 50, 15, ASSETS_ARRAY(ast_next_inrst_date, asset_counter)
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
                  GroupBox 20, 130, 140, 55, "Withdrawal Penalty"
                  Text 40, 150, 30, 10, "Amount:"
                  Text 30, 170, 40, 10, "Verification:"
                  GroupBox 165, 110, 100, 100, "Additional Owner(s)"
                  Text 170, 130, 40, 10, "Share Ratio:"
                  Text 170, 145, 50, 10, "Other owners:"
                  Text 5, 195, 65, 10, "Next Interest Date:"
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

                        If skip_this_panel = TRUE Then
                            err_msg = ""
                        End If

                        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

                    Loop until err_msg = ""
                    Call check_for_password(are_we_passworded_out)
                Loop until are_we_passworded_out = FALSE

                ASSETS_ARRAY(ast_ref_nbr, asset_counter) = left(ASSETS_ARRAY(ast_owner, asset_counter), 2)

                If count_cash_checkbox = checked Then ASSETS_ARRAY(apply_to_CASH, asset_counter) = "Y"
                If count_snap_checkbox = checked Then ASSETS_ARRAY(apply_to_SNAP, asset_counter) = "Y"
                If count_hc_checkbox = checked Then ASSETS_ARRAY(apply_to_HC, asset_counter) = "Y"
                If count_grh_checkbox = checked Then ASSETS_ARRAY(apply_to_GRH, asset_counter) = "Y"
                If count_ive_checkbox = checked Then ASSETS_ARRAY(apply_to_IVE, asset_counter) = "Y"

                If ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_one, asset_counter) = ""
                If ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_two, asset_counter) = ""
                If ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = "Type or Select" Then ASSETS_ARRAY(ast_othr_ownr_thr, asset_counter) = ""
                If share_ratio_denom = "1" Then ASSETS_ARRAY(ast_jnt_owner_YN, asset_counter) = "N"
                ASSETS_ARRAY(ast_own_ratio, asset_counter) = share_ratio_num & "/" & share_ratio_denom
                If ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = "Select..." Then ASSETS_ARRAY(ast_wthdr_verif, asset_counter) = ""

                If skip_this_panel = FALSE Then
                    Call navigate_to_MAXIS_screen("STAT", "ACCT")
                    EMWriteScreen ASSETS_ARRAY(ast_ref_nbr, asset_counter), 20, 76
                    If update_panel_type = "Existing ACCT" Then EMWriteScreen ASSETS_ARRAY(ast_instance, asset_counter), 20, 79
                    transmit
                    If update_panel_type = "New ACCT" Then EMWriteScreen "NN", 20, 79
                    If update_panel_type = "Existing ACCT" Then PF9

                    ASSETS_ARRAY(cnote_panel, asset_counter) = checked

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

                    If update_panel_type = "New ACCT" Then
                        EMReadScreen the_instance, 1, 2, 73
                        ASSETS_ARRAY(ast_instance, asset_counter) = "0" & the_instance
                    End If
                    transmit

                    if update_panel_type = "New ACCT" Then asset_counter = asset_counter + 1
                End If
            ElseIf panel_type = "SECU" Then
                ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                'Dialog to fill the SECU panel

                BeginDialog Dialog1, 0, 0, 271, 235, "New SECU panel for Case #" & MAXIS_case_number
                  DropListBox 75, 10, 135, 45, "Select One...", ast_owner
                  DropListBox 75, 30, 135, 45, "Select ..."+chr(9)+"SV - Savings"+chr(9)+"CK - Checking", ast_type
                  EditBox 75, 50, 105, 15, ast_number
                  EditBox 75, 70, 105, 15, ast_location
                  EditBox 75, 90, 50, 15, ast_balance
                  EditBox 160, 90, 50, 15, ast_bal_date
                  DropListBox 75, 110, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ast_verif
                  EditBox 95, 130, 60, 15, ast_face_value_LI
                  CheckBox 230, 25, 30, 10, "CASH", count_cash_checkbox
                  CheckBox 230, 40, 30, 10, "SNAP", count_snap_checkbox
                  CheckBox 230, 55, 20, 10, "HC", count_hc_checkbox
                  CheckBox 230, 70, 25, 10, "GRH", count_grh_checkbox
                  CheckBox 230, 85, 20, 10, "IVE", count_ive_checkbox
                  EditBox 75, 170, 50, 15, ast_wdrw_penlty
                  DropListBox 75, 190, 80, 45, "Select..."+chr(9)+"1 - Bank Statement"+chr(9)+"2 - Agcy Ver Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Document"+chr(9)+"6 - Personal Statement"+chr(9)+"N - No Ver Prvd", ast_wthdr_verif
                  EditBox 215, 125, 15, 15, share_ratio_num
                  EditBox 240, 125, 15, 15, share_ratio_denom
                  ComboBox 170, 160, 90, 45, "Type or Select", ast_othr_ownr_one
                  ComboBox 170, 175, 90, 45, "Type or Select", ast_othr_ownr_two
                  ComboBox 170, 190, 90, 45, "Type or Select", ast_othr_ownr_thr
                  ButtonGroup ButtonPressed
                    OkButton 160, 215, 50, 15
                    CancelButton 215, 215, 50, 15
                  Text 10, 15, 60, 10, "Owner of Account:"
                  Text 20, 35, 50, 10, "Security Type:"
                  Text 10, 55, 60, 10, "Security Number:"
                  Text 15, 75, 55, 10, "Security Name:"
                  Text 10, 95, 60, 10, "Cash Value (CSV):"
                  Text 25, 115, 40, 10, "Verification:"
                  Text 130, 95, 25, 10, "As of:"
                  GroupBox 225, 10, 40, 90, "Count:"
                  Text 10, 135, 75, 10, "Face Value of Life Ins:"
                  GroupBox 20, 155, 140, 55, "Withdrawal Penalty"
                  Text 40, 175, 30, 10, "Amount:"
                  Text 30, 195, 40, 10, "Verification:"
                  GroupBox 165, 110, 100, 100, "Additional Owner(s)"
                  Text 170, 130, 40, 10, "Share Ratio:"
                  Text 170, 145, 50, 10, "Other owners:"
                  Text 235, 125, 5, 10, "/"
                EndDialog


            ElseIf panel_type = "CARS" Then
                ReDim Preserve ASSETS_ARRAY(update_panel, asset_counter)
                'Dialog to fill the CARS panel.
            End If

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

    BeginDialog Dialog1, 0, 0, 266, 230, "AREP for Case # " & MAXIS_case_number
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
      CheckBox 70, 215, 35, 10, "SNAP", SNAP_AREP_check
      CheckBox 110, 215, 50, 10, "Health Care", HC_AREP_check
      CheckBox 165, 215, 30, 10, "Cash", CASH_AREP_check
      ButtonGroup ButtonPressed
        OkButton 210, 190, 50, 15
        CancelButton 210, 210, 50, 15
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
      Text 10, 145, 95, 10, "Date of AREP Form Recieved"
      Text 10, 195, 115, 10, "Date form was signed (MM/DD/YY):"
      Text 10, 210, 55, 20, "Programs Authorized for:"
    EndDialog

    Do
        Do
        	err_msg = ""
        	dialog Dialog1 					'Calling a dialog without a assigned variable will call the most recently defined dialog
        	cancel_confirmation

            If trim(arep_name) = "" Then err_msg = err_msg & vbNewLine & "* Enter the AREP's name."
            If update_AREP_panel_checkbox = checked Then
                If trim(arep_street) = "" OR trim(arep_city) = "" OR trim(arep_zip) = "" Then err_msg = err_msg & vbNewLine & "* Enter the street address of the AREP."
                If len(arep_name) > 37 Then err_msg = err_msg & vbNewLine & "* The AREP name is too long for MAXIS."
                If len(arep_street) > 44 Then err_msg = err_msg & vbNewLine & "* The AREP street is too long for MAXIS."
                If len(arep_city) > 15 Then err_msg = err_msg & vbNewLine & "* The AREP City is too long for MAXIS."
                If len(arep_state) > 2 Then err_msg = err_msg & vbNewLine & "* The AREP state is too long for MAXIS."
                If len(arep_zip) > 5 Then err_msg = err_msg & vbNewLine & "* The AREP zip is too long for MAXIS."
            End If
            If IsDate(AREP_recvd_date) = False Then err_msg = err_msg & vbNewLine & "* Enter the date the form was received."
        	IF SNAP_AREP_check <> checked AND HC_AREP_check <> checked AND CASH_AREP_check <> checked THEN err_msg = err_msg & vbNewLine &"* Select a program"
        	IF isdate(arep_signature_date) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was signed/valid from."
        	IF (TIKL_check = checked AND arep_signature_date = "") THEN err_msg = err_msg & vbNewLine & "* You have requested the script to TIKL based on the signature date but you did not enter the signature date."

        	IF err_msg <> "" THEN MsgBox "Plese resolve the following to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    'formatting programs into one variable to write in case note
    IF SNAP_AREP_checkbox = checked THEN AREP_programs = "SNAP"
    IF HC_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", HC"
    IF CASH_AREP_checkbox = checked THEN AREP_programs = AREP_programs & ", CASH"
    If left(AREP_programs, 1) = "," Then AREP_programs = right(AREP_programs, len(AREP_programs)-2)

    docs_rec = docs_rec & ", AREP Form"

    If update_AREP_panel_checkbox = checked Then

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

    If TIKL_check = checked then
    	call navigate_to_MAXIS_screen("dail", "writ")
    	call create_MAXIS_friendly_date(dateadd("m", 12, arep_signature_date), 0, 5, 18)
    	call write_variable_in_TIKL("Client's AREP release for HC is now 12 months old and no longer valid. Take appropriate action.")
    End If

End If

If ltc_1503_form_checkbox = checked Then


    'LTC 1503 gets it's own case note
    ' call start_a_blank_CASE_NOTE

End If

If left(docs_rec, 2) = ", " Then docs_rec = right(docs_rec, len(docs_rec)-2)        'trimming the ',' off of the list of docs


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
call start_a_blank_CASE_NOTE
If HSR_scanner_checkbox = 1 then
    Call write_variable_in_case_note("Docs Rec'd & scanned: " & docs_rec)
else
    Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
END IF
call write_bullet_and_variable_in_case_note("Document date stamp", doc_date_stamp)
If arep_form_checkbox = checked Then
    call write_variable_in_CASE_NOTE("* AREP FORM received on " & AREP_recvd_date & ". AREP: " & arep_name)
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
End If
call write_bullet_and_variable_in_case_note("JOBS", JOBS)
If evf_form_received_checkbox = checked Then
    call write_variable_in_CASE_NOTE("* EVF received " & evf_date_recvd & ": " & EVF_status_dropdown & "*")
    Call write_variable_in_CASE_NOTE("  - Employer Name: " & employer)
    Call write_variable_in_CASE_NOTE("  - EVF for HH member: " & evf_ref_numb)
    'for additional information needed
    IF info = "yes" then
        Call write_variable_in_CASE_NOTE("  - Additional Info requested: " & info & " on " & info_date & " by " & request_info)
    	If EVF_TIKL_checkbox = 1 then call write_variable_in_CASE_NOTE ("  ***TIKLed for 10 day return.***")
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
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " " & right(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 5) & " account. Balance: $" & ASSETS_ARRAY(ast_balance, the_asset) & " - Verif: " & right(ASSETS_ARRAY(ast_verif, the_asset), len(ASSETS_ARRAY(ast_verif, the_asset)) - 4))
                If ASSETS_ARRAY(ast_jnt_owner_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * Account is shared. Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " owns " & ASSETS_ARRAY(ast_own_ratio, the_asset) & " of the account.")
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
                Call write_variable_in_CASE_NOTE("  - Memb " & ASSETS_ARRAY(ast_ref_nbr, the_asset) & " - " & left(ASSETS_ARRAY(ast_type, the_asset), len(ASSETS_ARRAY(ast_type, the_asset)) - 4) & " - " & ASSETS_ARRAY(ast_year, the_asset) & " " & ASSETS_ARRAY(ast_make, the_asset) & " " & ASSETS_ARRAY(ast_model, the_asset) & " Value: $" & ASSETS_ARRAY(ast_trd_in, the_asset))
                If ASSETS_ARRAY(ast_owe_YN, the_asset) = "Y" Then Call write_variable_in_CASE_NOTE("    * $" & ASSETS_ARRAY(ast_amt_owed, the_asset) & " owed as of " & ASSETS_ARRAY(ast_owed_date, the_asset) & " - Verif: " & ASSETS_ARRAY(ast_owe_verif, the_asset))
            End If
        Next
    End If

    If LTC_case = vbYes Then

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
IF HSR_scanner_checkbox = 1 then Call write_variable_in_case_note("* Documents imaged to ECF.")
call write_bullet_and_variable_in_case_note("Verifications still needed", verifs_needed)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure(end_msg)
