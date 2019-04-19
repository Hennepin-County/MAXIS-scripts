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

If LTC_case = vbNo then end_msg = "Sucess! Documents received noted for case."
If LTC_case = vbYes then end_msg = "Sucess! Documents received noted for LTC case."

'EVF HANDLING =======================================================================================
If evf_form_received_checkbox = checked Then
    Tikl_checkbox = checked 'defaulting the TIKL checkbox to be checked initially in the dialog.
    date_received = doc_date_stamp

    BeginDialog Dialog1, 0, 0, 291, 205, "Employment Verification Form Received"
      Text 70, 10, 60, 10, MAXIS_case_number
      EditBox 220, 5, 60, 15, date_received
      ComboBox 70, 30, 210, 15, "Select one..."+chr(9)+"Signed by Client & Completed by Employer"+chr(9)+"Signed by Client"+chr(9)+"Completed by Employer", EVF_status_dropdown
      EditBox 70, 50, 210, 15, employer
      EditBox 70, 70, 210, 15, client
      DropListBox 75, 110, 60, 15, "Select one..."+chr(9)+"yes"+chr(9)+"no", info
      EditBox 220, 110, 60, 15, info_date
      EditBox 75, 130, 60, 15, request_info
      CheckBox 160, 135, 105, 10, "10 day TIKL for additional info", Tikl_checkbox
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
    		cancel_confirmation 		'asks if you want to cancel and if "yes" is selected sends StopScript
    		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You need to type a valid case number."
    		IF IsDate(date_received) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date for date the EVF was received."
    		If EVF_status_dropdown = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select the status of the EVF on the dropdown menu"		'checks that there is a date in the date received box
    		IF employer = "" THEN err_msg = err_msg & vbCr & "* You must enter the employers name."  'checks if the employer name has been entered
    		IF client = "" THEN err_msg = err_msg & vbCr & "* You must enter the MEMB information."  'checks if the client name has been entered
    		IF info = "Select one..." THEN err_msg = err_msg & vbCr & "* You must select if additional info was requested."  'checks if completed by employer was selected
    		IF info = "yes" and IsDate(info_date) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid date that additional info was requested."  'checks that there is a info request date entered if the it was requested
    		IF info = "yes" and request_info = "" THEN err_msg = err_msg & vbCr & "* You must enter the method used to request additional info."		'checks that there is a method of inquiry entered if additional info was requested
    		If info = "no" and request_info <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have information requested."
    		If info = "no" and info_date <> "" then err_msg = err_msg & vbCr & "* You cannot mark additional info as 'no' and have a date requested."
    		If Tikl_checkbox = 1 and info <> "yes" then err_msg = err_msg & vbCr & "* Additional informaiton was not requested, uncheck the TIKL checkbox."
    		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* You must enter your actions taken."		'checks that notes were entered
    		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!" 		'checks that the case note was signed
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "* Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    docs_rec = docs_rec & ", EVF for M" & client

    'Checks if additional info is yes and the TIKL is checked, sets a TIKL for the return of the info
    If Tikl_checkbox = checked Then
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

End If


If asset_form_checkbox = checked Then

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

End If
call write_bullet_and_variable_in_case_note("JOBS", JOBS)
If evf_form_received_checkbox = checked Then
    call write_variable_in_CASE_NOTE("* EVF received " & date_received & ": " & EVF_status_dropdown & "*")
    Call write_variable_in_CASE_NOTE("  - Employer Name: " & employer)
    Call write_variable_in_CASE_NOTE("  - EVF for HH member: " & client)
    'for additional information needed
    IF info = "yes" then
        Call write_variable_in_CASE_NOTE("  - Additional Info requested: " & info & " on " & info_date & " by " & request_info)
    	If Tikl_checkbox = 1 then call write_variable_in_CASE_NOTE ("  ***TIKLed for 10 day return.***")
    Else
        Call write_variable_in_CASE_NOTE("  - No additional information is needed/requested.")
    END IF
End If
call write_bullet_and_variable_in_case_note("BUSI", BUSI)
call write_bullet_and_variable_in_case_note("BUSI/RBIC", BUSI_RBIC)
call write_bullet_and_variable_in_case_note("UNEA", UNEA)
If asset_form_checkbox = checked Then

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
