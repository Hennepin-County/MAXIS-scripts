'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EMERGENCY SCREENING.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         'manual run time in seconds
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

'LOCAL FUNCTIONS============================================================================================================
FUNCTION VERIF_BUTTONS
	If ButtonPressed = MEMB_number then
		MEMB_function
		HH_member_array = ""
		FOR i = 0 to total_clients
			IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
				IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
					'msgbox all_clients_
					HH_member_array = Right(all_clients_array(i, 0), len(all_clients_array(i, 0))   ) & ", " & HH_member_array
				END IF
			END IF
		NEXT
		hh_size_split = Len(HH_member_array) - Len(Replace(HH_member_array,",",""))
          hh_size = CStr(hh_size_split)
     End If
     If ButtonPressed = ei_button then
          ei_function
     End If
     If ButtonPressed = unearn_income then
          unea_function
     End If
     If ButtonPressed = shelter_button then
          shel_function
     End If
	If ButtonPressed = expense_button then
		expense_function
	End If
	'This part works with the prev/next buttons on several of our dialogs. You need to name your buttons prev_panel_button, next_panel_button, prev_memb_button, and next_memb_button in order to use them.
	EMReadScreen STAT_check, 4, 20, 21
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		End if
	End if
	If ButtonPressed = ADDR_button then call navigate_to_MAXIS_screen("stat", "ADDR")
     If ButtonPressed = SHEL_button then call navigate_to_MAXIS_screen("stat", "SHEL")
	If ButtonPressed = BUSI_button then call navigate_to_MAXIS_screen("stat", "BUSI")
	If ButtonPressed = JOBS_button then call navigate_to_MAXIS_screen("stat", "JOBS")
	If ButtonPressed = MEMB_button then call navigate_to_MAXIS_screen("stat", "MEMB")
	If ButtonPressed = TYPE_button then call navigate_to_MAXIS_screen("stat", "TYPE")
	If ButtonPressed = PROG_button then call navigate_to_MAXIS_screen("stat", "PROG")
	If ButtonPressed = REVW_button then call navigate_to_MAXIS_screen("stat", "REVW")
	If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("stat", "UNEA")
	If ButtonPressed = CURR_button then call navigate_to_MAXIS_screen("case", "CURR")
	If ButtonPressed = INQX_button then
		Call navigate_to_MAXIS_screen("MONY", "INQX")
		EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
		EMWriteScreen begin_search_year, 6, 41
		EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
		EMWriteScreen MAXIS_footer_year, 6, 56
		transmit
	End If
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		End if
	End if
END FUNCTION

'this function creates the hh member dynamic dialog
FUNCTION MEMB_function

    Dialog1 = ""
    BEGINDIALOG Dialog1, 0,  0, 256, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
    	Text 10, 5, 145, 10, "Who is applying?:"
    	FOR clt_i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
    		IF all_clients_array(clt_i, 0) <> "" THEN checkbox 10, (20 + (clt_i * 15)), 150, 10, all_clients_array(clt_i, 0), all_clients_array(clt_i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
    	NEXT
    	ButtonGroup ButtonPressed
    	OkButton 200, 20, 50, 15
    	'CancelButton 155, 40, 50, 15
    ENDDIALOG
    Dialog Dialog1
End Function

FUNCTION ei_function
    memb_gross_total = 0
    memb_net_total = 0
    f = 0
    dialog_measures = 0
    FOR clt_i = 0 to total_clients
    	IF all_clients_array(clt_i, 0) <> "" THEN
    		IF all_clients_array(clt_i, 1) = 1 THEN dialog_measures = dialog_measures + 1
    	End If
    Next
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 241, (45 + (dialog_measures * 35)), "EARN INCOME for Members in HH"
        FOR clt_i = 0 to total_clients
        	IF all_clients_array(clt_i, 0) <> "" THEN
        		IF all_clients_array(clt_i, 1) = 1 THEN Editbox 40, (20 + (f * 35)), 50, 15, memb_gross(clt_i)
        		IF all_clients_array(clt_i, 1) = 1 THEN Editbox 115, (20 + (f * 35)), 50, 15, memb_net(clt_i)
        		IF all_clients_array(clt_i, 1) = 1 THEN DropListBox 170, (20 + (f * 35)), 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", gross_income_verification(clt_i)
        		IF all_clients_array(clt_i, 1) = 1 THEN GroupBox 5, (10 + (f * 35)), 230, 30, all_clients_array(clt_i, 0)
        		IF all_clients_array(clt_i, 1) = 1 THEN Text 100, (25 + (f* 35)), 15, 10, "Net:"
        		IF all_clients_array(clt_i, 1) = 1 THEN Text 15, (25 + (f * 35)), 25, 10, "Gross:"
        		IF all_clients_array(clt_i, 1) = 1 THEN f = f + 1
        	End If
        Next
      ButtonGroup ButtonPressed
        IF total_clients <> "" THEN OkButton 95, (20 + (f * 35)), 50, 15
    EndDialog

    Dialog Dialog1
End Function

FUNCTION unea_function
	Call remove_dash_from_droplist(UNEA_type_list)
	UNEA_type_list = replace(UNEA_type_list, "01 RSDI, Disa"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "02 RSDI, No Disa"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "03 SSI"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "08 Direct Child Support"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "36 Disbursed Child Support"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "39 Disbursed CS Arrears"+chr(9), "")
	UNEA_type_list = replace(UNEA_type_list, "43 Disbursed Excess CS"+chr(9), "")
    case_memb_unea_total = 0
    f = 0
    dialog_measures = 0
    FOR clt_i = 0 to total_clients
    	IF all_clients_array(clt_i, 0) <> "" THEN
    		IF all_clients_array(clt_i, 1) = 1 THEN dialog_measures = dialog_measures + 1
    	End If
    Next
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 481,(45 + (dialog_measures * 35)), "UNEARN INCOME"
        FOR clt_i = 0 to total_clients
            IF all_clients_array(clt_i, 0) <> "" THEN
              IF all_clients_array(clt_i, 1) = 1 THEN EditBox 30, (15 + (f * 35)), 50, 15, ssi_income(clt_i)
              IF all_clients_array(clt_i, 1) = 1 THEN EditBox 110, (15 + (f * 35)), 50, 15, rsdi_income(clt_i)
              IF all_clients_array(clt_i, 1) = 1 THEN EditBox 215, (15 + (f * 35)), 50, 15, child_support(clt_i)
              IF all_clients_array(clt_i, 1) = 1 THEN EditBox 420, (15 + (f * 35)), 50, 15, other_unea(clt_i)
              IF all_clients_array(clt_i, 1) = 1 THEN Text 170, (20 + (f * 35)), 45, 10, "Child Support:"
              IF all_clients_array(clt_i, 1) = 1 THEN Text 415, (20 + (f * 35)), 5, 10, "="
              IF all_clients_array(clt_i, 1) = 1 THEN Text 90, (20 + (f * 35)), 20, 10, "RSDI:"
              IF all_clients_array(clt_i, 1) = 1 THEN Text 15, (20 + (f * 35)), 15, 10, "SSI:"
              IF all_clients_array(clt_i, 1) = 1 THEN GroupBox 5, (5 + (f * 35)), 470, 30, all_clients_array(clt_i, 0)
			  IF all_clients_array(clt_i, 1) = 1 THEN DropListBox 300, (15 + (f * 35)), 110, 15, "Select One..."+chr(9)+UNEA_type_list, other_unea_type(clt_i)
              IF all_clients_array(clt_i, 1) = 1 THEN f = f + 1
            End If
        Next
        ButtonGroup ButtonPressed
          IF total_clients <> "" THEN OkButton 215, (15 + (f * 35)), 50, 15
    EndDialog

    Dialog Dialog1
End Function

Function shel_function
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 211, 275, "Shelter Information Calc"
      EditBox 85, 15, 50, 15, rent_portion
      DropListBox 140, 15, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", rent_verification
      EditBox 85, 35, 50, 15, other_fees
      DropListBox 140, 35, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", Other_fees_verification
      EditBox 85, 75, 50, 15, rent_due
      DropListBox 140, 75, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", rent_due_verification
      EditBox 85, 95, 50, 15, late_fees
      DropListBox 140, 95, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", late_fees_verification
      EditBox 85, 115, 50, 15, damage_dep
      DropListBox 140, 115, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", dd_verification
      EditBox 85, 135, 50, 15, court_fees
      DropListBox 140, 135, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", court_fees_verification
      EditBox 85, 155, 50, 15, hest_due
      DropListBox 140, 155, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", utility_verification
      DropListBox 140, 190, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", eviction_verification
      DropListBox 140, 210, 60, 45, "Verifications?"+chr(9)+"Requested"+chr(9)+"Received", disconnection_verification
      ButtonGroup ButtonPressed
        OkButton 80, 245, 50, 15
      Text 65, 20, 20, 10, "Rent:"
      Text 50, 100, 35, 10, "Late fees:"
      GroupBox 5, 65, 200, 110, "Expenses Due"
      Text 50, 140, 35, 10, "Court fees:"
      GroupBox 5, 5, 200, 55, "Monthly Expenses"
      Text 65, 80, 20, 10, "Rent:"
      Text 65, 160, 20, 10, "Utility:"
      Text 30, 120, 55, 10, "Damage Deposit:"
      Text 85, 195, 55, 10, "Eviciton Notice:"
      Text 10, 40, 75, 10, "Other fees(garage,etc):"
      Text 65, 215, 75, 10, "Disconnection Notice:"
      GroupBox 5, 180, 200, 50, "Other Verifications"
    EndDialog

    Dialog Dialog1
End Function

Function expense_function
    Do
        err_msg = ""
        fs_mf_total = Cstr(Cint(mf_fs_amt_total) + Cint(fs_amt_total))
        fs_expense = CStr(Cint(thrifty_food) - Cint(fs_mf_total))
        if fs_expense < 0 then fs_expense = "0"
        food_allotment_expense = "Food Allotment($" & thrifty_food & ") - FS/MF-FS issued ($" & fs_mf_total & ") = $" & fs_expense

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 301, 230, "Living Expense Paid from:" & chr(9) & dateadd("d", -30, app_date) & chr(9) & " To:" & chr(9) & dateadd("d", -1, app_date) & chr(9)
          EditBox 185, 45, 50, 15, shel_paid
          EditBox 185, 65, 50, 15, hest_paid
          DropListBox 185, 85, 50, 45, "Select One"+chr(9)+"Yes"+chr(9)+"No", flat_living_expense
          DropListBox 185, 105, 50, 45, "Select One"+chr(9)+"Yes"+chr(9)+"No", flat_trans
          DropListBox 185, 125, 50, 45, "Select One"+chr(9)+"Yes"+chr(9)+"No", flat_phone
          EditBox 185, 145, 50, 15, actual_paid
          EditBox 185, 165, 50, 15, other_paid
          ButtonGroup ButtonPressed
            OkButton 130, 200, 50, 15
          Text 100, 90, 85, 10, "Flat $500 Living Expense:"
          Text 110, 50, 75, 10, "Shelter Expense Paid:"
          ButtonGroup ButtonPressed
            OkButton 130, 200, 50, 15
          Text 110, 150, 75, 10, "Actual Living Expense:"
          Text 10, 110, 175, 10, "Storage, Transportation Flat $113.50(working person):"
          Text 145, 70, 40, 10, "Utility Paid:"
          Text 130, 130, 55, 10, "Flat $38 Phone:"
          Text 165, 170, 20, 10, "Other:"
          Text 5, 30, 285, 10, chr(9) & food_allotment_expense & chr(9)
          GroupBox 5, 10, 285, 175, "Living Expense Paid from:" & chr(9) & dateadd("d", -30, app_date) & chr(9) & " To:" & chr(9) & dateadd("d", -1, app_date) & chr(9)
        EndDialog

        Dialog Dialog1

        If shel_paid = "" then shel_paid = "0"
        If hest_paid = "" then hest_paid = "0"
        If actual_paid = "" then actual_paid = "0"
        If other_paid = "" then other_paid = "0"

        If actual_paid <> "0" and flat_living_expense = "Yes" then err_msg = "You selected 'Yes' for Flat $500 Living Expense, you cannot list amounts in 'Actual Living Expense field.' Please correct this."
        IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
    Loop until err_msg = ""
End Function

'THE SCRIPT================================================================================================================================
EMConnect ""

call check_for_MAXIS(False)	'checking for an active MAXIS session

call MAXIS_case_number_finder(MAXIS_case_number)
'Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'declares defaults'
Dim memb_gross(50)
Dim memb_net(50)
Dim income_list_gross(50)
Dim income_list_net(50)
Dim ssi_income(50)
Dim rsdi_income(50)
Dim child_support(50)
Dim other_unea(50)
Dim other_unea_type(50)
Dim unea_ssi_list(50)
Dim unea_rsdi_list(50)
Dim unea_child_support_list(50)
Dim unea_other_list(50)
Dim adult_applying(50)
Dim adult_not_applying(50)
Dim gross_income_verification(50)
Dim gross_income_received_list(50)
Dim gross_income_verification_list(50)

case_memb_unea_total = "0"
memb_gross_total = "0"
memb_net_total = "0"

FOR clt_i = 0 to 50
	memb_gross(clt_i) = "0"
	memb_net(clt_i) = "0"
	income_list_gross(clt_i) = 0
	income_list_net(clt_i) = 0
	ssi_income(clt_i) = "0"
	rsdi_income(clt_i) = "0"
	child_support(clt_i) = "0"
	other_unea(clt_i) = "0"
Next


rent_portion = "0"
other_fees = "0"
rent_due = "0"
late_fees = "0"
court_fees = "0"
hest_due = "0"
damage_dep = "0"
shel_paid = "0"
hest_paid = "0"
actual_paid = "0"
other_paid = "0"
total_expense = "0"

'formats default date'
If len(datepart("m", date())) = 1 then
	m = "0" & datepart("m", date())
Else
	m = datepart("m",date())
End IF
If len(datepart("d", date())) = 1 then
	d = "0" & datepart("d", date())
Else
	d = datepart("d",date())
End IF

If len(datepart("m", date()-30)) = 1 then
	ea_eval_m = "0" & datepart("m", date()+10)
Else
	ea_eval_m = datepart("m",date()-30)
End IF
If len(datepart("d", date()-30)) = 1 then
	ea_eval_d = "0" & datepart("d", date()+10)
Else
	ea_eval_d = datepart("d",date()-30)
End IF
app_date= m & "/" & d & "/" & right(datepart("yyyy", date()), 2)
'determines EA Eval Period'
ea_eval_date = ea_eval_m & "/" & ea_eval_d & "/" & right(datepart("yyyy", date()-30), 2)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 171, 120, "EA/EGA Screening"
  EditBox 90, 5, 75, 15, MAXIS_case_number
  EditBox 105, 25, 60, 15, app_date
  DropListBox 105, 45, 60, 45, "Select One"+chr(9)+"EA"+chr(9)+"EGA", prog_type_case_dialog
  EditBox 70, 65, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 90, 50, 15
    CancelButton 90, 90, 50, 15
  Text 25, 30, 75, 10, "Date of app (xx/xx/xx):"
  Text 5, 70, 65, 10, "Worker's Signature:"
  Text 65, 50, 35, 10, "EA/EGA?:"
  Text 60, 10, 25, 10, "Case #:"
EndDialog

'The Script'
Do
	err_msg = ""
	Dialog Dialog1
	cancel_confirmation
	If MAXIS_case_number = "" then err_msg = err_msg & vbCr & "You must have a case number to continue."
	If len(MAXIS_case_number) > 8 then err_msg = err_msg & vbCr & "Your case number need to be 8 digits or less."
	If prog_type_case_dialog = "Select One" then err_msg = err_msg & vbCr & "You must choose a program type."
	If DateValue(app_date) > Date() then err_msg = err_msg & vbCr & "You cannot enter a future application date."
	If err_msg <> "" then Msgbox err_msg
	call check_for_password (are_we_passworded_out) 'adding functionality for MAXIS v.6 Password Out issue'
Loop until err_msg = ""

'reformats App Date Again'
If len(datepart("m", app_date)) = 1 then
	m = "0" & datepart("m", app_date)
Else
	m = datepart("m", app_date)
End IF

MAXIS_footer_month = m
MAXIS_footer_year = right(datepart("yyyy", app_date), 2)
back_to_self


'DATE CALCULATIONS From Isle Hennepin County
'creating month variable 13 months prior to current footer month/year to search for EMER programs issued
begin_search_month = dateadd("m", -13, app_date)
begin_search_year = datepart("yyyy", begin_search_month)
begin_search_year = right(begin_search_year, 2)
begin_search_month = datepart("m", begin_search_month)
If len(begin_search_month) = 1 then begin_search_month = "0" & begin_search_month
'End of date calculations----------------------------------------------------------------------------------------------

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_search_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_search_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 9, 50		'selecting EA
EMWriteScreen "x", 11, 50		'selecting EGA
transmit

'searching for EA/EG issued on the INQD screen
DO
	row = 6
	DO
		EMReadScreen emer_issued, 1, row, 16		'searching for EMER programs as they start with E
		IF emer_issued = "E" then
			'reading the EMER information for EMER issuance
			EMReadScreen EMER_type, 2, row, 16
			EMReadScreen EMER_amt_issued, 7, row, 39
			EMReadScreen EMER_elig_start_date, 8, row, 7
			'EMReadScreen EMER_elig_end_date, 8, row, 73
			exit do
		ELSE
			row = row + 1
		END IF
	Loop until row = 18				'repeats until the end of the page
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

'creating variables and conditions for EMER screening
New_EMER_year = dateadd("YYYY", 1, EMER_elig_start_date)
EMER_available_date = dateadd("d", 1, New_EMER_year)	'creating emer available date that is 1 day & 1 year past the EMER_elig_end_date
EMER_last_used_dates = EMER_elig_start_date ''& " - " & EMER_elig_end_date	'combining dates into new variable

If emer_issued <> "E" or datevalue(app_date) > datevalue(EMER_available_date) then	'creating variables for cases that have not had EMER issued in current 13 months
 	EMER_last_used_dates = "n/a"
	EMER_available_date = "Currently available"
END IF

'Declares a variable from EA Evaluation start date to be use for inqx search programs'
begin_eval_day = dateadd("d", -30, app_date)
begin_eval_month = datepart("m", begin_eval_day)
begin_eval_year = datepart("yyyy", begin_eval_day)
begin_eval_year = right(begin_eval_year, 2)
If len(begin_eval_month)= 1 then begin_eval_month = "0" & begin_eval_month

'Screen FS Prog'

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_eval_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_eval_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 9, 5

transmit

Dim fs_amt_issued(50)
fs_amt_issued(0) = "0"
h = 1
i = 0
For j = 6 to 18
	EMReadScreen issued_date, 8, j, 7
		If issued_date <> "        " then
          	If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then
	            		EMReadScreen fs_amt_issued(h), 6, j, 39
					fs_amt_issued(h) = replace(fs_amt_issued(h), " ","")
                    	fs_amt_total = Cint(fs_amt_issued(h)) + Cint(fs_amt_issued(i))
                 		h = h + 1
                 		i = i + 1
					fs_prog = true
			End If
		End If
Next

'Screen MFIP Prog'

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWriteScreen begin_eval_month, 6, 38		'entering footer month/year 13 months prior to current footer month/year'
EMWriteScreen begin_eval_year, 6, 41
EMWriteScreen MAXIS_footer_month, 6, 53		'entering current footer month/year
EMWriteScreen MAXIS_footer_year, 6, 56
EMWriteScreen "x", 10, 5

transmit

Dim mf_amt_issued(50)
Dim mf_fs_amt_issued(50)
Dim mf_hg_amt_issued(50)
mf_amt_issued(0) = "0"
mf_fs_amt_issued(0) = "0"
mf_hg_amt_issued(0) = "0"
h = 1
i = 0
f = 1
g = 0
s = 1
d = 0
For j = 6 to 18
	EMReadScreen issued_date, 8, j, 7
	EMReadScreen prog_type, 5, j, 16
		If issued_date <> "        " then
			If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then

				If prog_type = "MF-MF" then
						EMReadScreen mf_amt_issued(h), 6, j, 39
						mf_amt_issued(h) = replace(mf_amt_issued(h), " ","")
                    		mf_amt_total = Cint(mf_amt_issued(h)) + Cint(mf_amt_issued(i))
                 			h = h + 1
                 			i = i + 1
						mf_prog = true
				End If
			End If
		End If
Next
For j = 6 to 18
	EMReadScreen issued_date, 8, j, 7
	EMReadScreen prog_type, 5, j, 16
		If issued_date <> "        " then
			If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then

				If prog_type = "MF-FS" then
						EMReadScreen mf_fs_amt_issued(f), 6, j, 39
						mf_fs_amt_issued(f) = replace(mf_fs_amt_issued(f), " ","")
                    		mf_fs_amt_total = Cint(mf_fs_amt_issued(f)) + Cint(mf_fs_amt_issued(g))
                 			f = f + 1
                 			g = g + 1
						mf_fs_prog = true
				End If
			End If
		End If
Next
For j = 6 to 18
	EMReadScreen issued_date, 8, j, 7
	EMReadScreen prog_type, 5, j, 16
		If issued_date <> "        " then
			If cdate(issued_date) =< cdate(dateadd("d", -1, app_date)) and cdate(issued_date) >= cdate(dateadd("d", -30, app_date)) then

				If prog_type = "MF-HG" then
						EMReadScreen mf_hg_amt_issued(s), 6, j, 39
						mf_hg_amt_issued(s) = replace(mf_hg_amt_issued(s), " ","")
                    		mf_hg_amt_total = Cint(mf_hg_amt_issued(s)) + Cint(mf_hg_amt_issued(d))
                 			s = s + 1
                 			d = d + 1
						mf_hg_prog = true
				End If
			End If
		End If
Next

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadscreen last_name_array, 25, 6, 30
	EMReadscreen first_name_array, 12, 6, 63
	EMReadscreen client_age, 2, 8, 76
	client_age = replace(client_age, " ", "")
	If Cint(client_age) >= 20 then
		client_is = "(ADULT)"
	Else
		client_is = "(CHILD)"
	End If
	last_name_array = replace(last_name_array, "_", "")
	last_name_array = Lcase(last_name_array)
	last_name_array = UCase(Left(last_name_array, 1)) &  Mid(last_name_array, 2)
	first_name_array = replace(first_name_array, "_", "") '& " "
	first_name_array = Lcase(first_name_array)
	first_name_array = UCase(Left(first_name_array, 1)) &  Mid(first_name_array, 2)
	client_string = ref_nbr & " " & first_name_array & " " & last_name_array & " " & client_is
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
DIM all_client_array()
ReDim all_clients_array(total_clients, 1)
FOR clt_x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(clt_x, 0) = Interim_array(clt_x)
	all_clients_array(clt_x, 1) = 1
NEXT
HH_size = CStr(total_clients)


If prog_type_case_dialog = "EA" then
	Do
		err_msg = ""
		If HH_size = 0 then
			FPG_size = "$0.00"
			thrifty_food = "0"
		End If
		If HH_size = 1 then
			FPG_size = "$1962"
			thrifty_food = "194"
		End If
		If HH_size = 2 then
			FPG_size = "$2655"
			thrifty_food = "357"
		End If
		If HH_size = 3 then
			FPG_size = "$3348"
			thrifty_food = "511"
		End If
		If HH_size = 4 then
			FPG_size = "$4042"
			thrifty_food = "649"
		End If
		If HH_size = 5 then
			FPG_size = "$4735"
			thrifty_food = "771"
		End If
		If HH_size = 6 then
			FPG_size = "$5428"
			thrifty_food = "925"
		End If
		If HH_size = 7 then
			FPG_size = "$6122"
			thrifty_food = "1022"
		End If
		If HH_size = 8 then
			FPG_size = "$6815"
			thrifty_food = "1169"
		End If
		If HH_size = 9 then
			FPG_size = "$7508"
			thrifty_food = "1315"
		End If
		If HH_size = 10 then
			FPG_size = "$8202"
			thrifty_food = "1461"
		End If
		If HH_size = 11 then
			FPG_size = "$8895"
			thrifty_food = "1607"
		End If
		If HH_size = 12 then
			FPG_size = "$9588"
			thrifty_food = "1753"
		End If
		If HH_size = 13 then
			FPG_size = "$10281"
			thrifty_food = "1899"
		End If
		If HH_size = 14 then
			FPG_size = "$10974"
			thrifty_food = "2045"
		End If
		If HH_size = 15 then
			FPG_size = "$11667"
			thrifty_food = "2191"
		End If
		If HH_size = 16 then
			FPG_size = "$12360"
			thrifty_food = "2337"
		End If
		If HH_size = 17 then
			FPG_size = "$13053"
			thrifty_food = "2483"
		End If
		If HH_size = 18 then
			FPG_size = "$13746"
			thrifty_food = "2629"
		End If
		If HH_size = 19 then
			FPG_size = "$14439"
			thrifty_food = "2775"
		End If
		If HH_size = 20 then
			FPG_size = "$15132"
			thrifty_food = "2921"
		End If
			'total gross and net calculations
			FOR clt_i = 0 to 50
			income_list_gross(clt_i) = 0
			income_list_net(clt_i) = 0
		Next

		'totals expenses'
		rent_mo = Cstr(Cint(rent_portion) + Cint(other_fees))
		shel_due = Cstr(Cint(rent_due) + Cint(late_fees) + Cint(court_fees) + Cint(hest_due)+Cint(damage_dep))

		'generating verif request list'
		i = 1
		If rent_verification = "Requested" then
			rent_verification_list = i & ") monthly/rent cost, "
			i = i + 1
		Else
			rent_verification_list = ""
		End If
		If Other_fees_verification = "Requested" then
			Other_fees_verification_list = i & ") other monthly fees, "
			i = i + 1
		Else
			Other_fees_verification_list = ""
		End If
		If rent_due_verification = "Requested" then
			rent_due_verification_list = i & ") rent due balance, "
			i = i + 1
		Else
			rent_due_verification_list = ""
		End If
		If late_fees_verification = "Requested" then
			late_fees_verification_list = i & ") late fees, "
			i = i + 1
		Else
			late_fees_verification_list = ""
		End If
		If dd_verification = "Requested" then
			dd_verification_list = i & ") damage deposit fee, "
			i = i + 1
		Else
			dd_verification_list = ""
		End If
		If court_fees_verification = "Requested" then
			court_fees_verification_list = i & ") court fees, "
			i = i + 1
		Else
			court_fees_verification_list = ""
		End If
		If utility_verification = "Requested" then
			utility_verification_list = i & ") utility bills, "
			i = i + 1
		Else
			utility_verification_list = ""
		End If
		If eviction_verification = "Requested" then
			eviction_verification_list = i & ") eviction notice, "
			i = i + 1
		Else
			eviction_verification_list = ""
		End If
		If disconnection_verification = "Requested" then
			disconnection_verification_list = i & ") disconnection notice, "
			i = i + 1
		Else
			disconnection_verification_list = ""
		End If

		'Earnincome verification request list'
		gross_income_verification_list(0) = ""
		w = 1
		v = 0

		FOR clt_i = 0 to total_clients
			IF all_clients_array(clt_i, 0) <> "" THEN
				IF all_clients_array(clt_i, 1) = 1 THEN
					If gross_income_verification(clt_i) = "Requested" then
						gross_income_verification_list(w) = gross_income_verification_list(v) & i & ") Paystubs for: " & right(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8), len(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8)) - 4) & ", "
						i = i + 1
						w = w + 1
						v = v + 1
					End If
				End If
			End If
		Next

		Verif_request_list = rent_verification_list & Other_fees_verification_list & rent_due_verification_list & late_fees_verification_list & dd_verification_list & court_fees_verification_list & utility_verification_list & eviction_verification_list & disconnection_verification_list & gross_income_verification_list(v)

		i = 1
		If rent_verification = "Received" then
			rent_received_list = i & ") monthly/rent cost, "
			i = i + 1
		Else
			rent_received_list = ""
		End If
		If Other_fees_verification = "Received" then
			Other_fees_received_list = i & ") other monthly fees, "
			i = i + 1
		Else
			Other_fees_received_list = ""
		End If
		If rent_due_verification = "Received" then
			rent_due_received_list = i & ") rent due balance, "
			i = i + 1
		Else
			rent_due_received_list = ""
		End If
		If late_fees_verification = "Received" then
			late_fees_received_list = i & ") late fees, "
			i = i + 1
		Else
			late_fees_received_list = ""
		End If
		If dd_verification = "Received" then
			dd_received_list = i & ") damage deposit fee, "
			i = i + 1
		Else
			dd_received_list = ""
		End If
		If court_fees_verification = "Received" then
			court_fees_received_list = i & ") court fees, "
			i = i + 1
		Else
			court_fees_received_list = ""
		End If
		If utility_verification = "Received" then
			utility_received_list = i & ") utility bills, "
			i = i + 1
		Else
			utility_received_list = ""
		End If
		If eviction_verification = "Received" then
			eviction_received_list = i & ") eviction notice, "
			i = i + 1
		Else
			eviction_received_list = ""
		End If
		If disconnection_verification = "Received" then
			disconnection_received_list = i & ") disconnection notice, "
			i = i + 1
		Else
			disconnection_received_list = ""
		End If

		'Earnincome verification request list'
		gross_income_received_list(0) = ""
		w = 1
		v = 0

		FOR clt_i = 0 to total_clients
			IF all_clients_array(clt_i, 0) <> "" THEN
				IF all_clients_array(clt_i, 1) = 1 THEN
					If gross_income_verification(clt_i) = "Received" then
						gross_income_received_list(w) = gross_income_received_list(v) & i & ") Paystubs for: " & right(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8), len(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8)) - 4) & ", "
						i = i + 1
						w = w + 1
						v = v + 1
					End If
				End If
			End If
		Next

		Verif_received_list = rent_received_list & Other_fees_received_list & rent_due_received_list & late_fees_received_list & dd_received_list & court_fees_received_list & utility_received_list & eviction_received_list & disconnection_received_list & gross_income_received_list(v)





		'dummy variable to get total of adults and ratio responsibility'
		f = 0
		g = 0
		'Looks for adults applying'
		FOR clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN

		  IF all_clients_array(clt_i, 1) = 1 THEN
		  	If InStr(all_clients_array(clt_i, 0), "(ADULT)") <> 0 then
			adult_applying(f) = all_clients_array(clt_i, 0)
			f = f + 1
			End If
		  End If
		End If
		Next
		adult_not_applying(0) = ""
		h = 1
		i = 0

		'Looks for adults not applying'
		FOR clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN

		  IF all_clients_array(clt_i, 1) = 0 THEN
			If InStr(all_clients_array(clt_i, 0), "(ADULT)") <> 0 then
			adult_not_applying(h) = adult_not_applying(i) & right(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8), len(left(all_clients_array(clt_i, 0), len(all_clients_array(clt_i, 0)) - 8)) - 4) & ", "
			f = f + 1
			g = g + 1
			h = h + 1
			i = i + 1
			End If
		  End If
		End If
		Next

		number_of_adults_hh = f
		number_of_adults_not_applying_responsible = g
		ratio_responsibility = g/f
		adult_not_applying_portion_of_due = Left((shel_due * ratio_responsibility), 7)
		adult_not_applying_each_portion_of_due = shel_due /f
		If g <> 0 then
			If shel_due <> "0" then
			hh_msg = "Not applying: " & adult_not_applying(i) & " The bal/ratio is split by " & f & " adults in the HH. $" & adult_not_applying_portion_of_due & " must be paid first to pass cost/eff test."
			test_pass7 = false
			End If
		Else
			hh_msg = ""
			test_pass7 = true
		End If

		'Gross/Net income calculations'
		memb_gross_total = 0
		memb_net_total = 0
		f = 0
		For clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN
			IF all_clients_array(clt_i, 1) = 1 THEN income_list_gross(f) = memb_gross(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN income_list_net(f) = memb_net(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN f = f + 1
		End If
		Next
		For clt_i = 0 to total_clients
			If memb_gross(clt_i) = "" then memb_gross(clt_i) = "0"
			If memb_net(clt_i) = "" then memb_net(clt_i) = "0"
		Next
		FOR clt_i = 0 to 49
		if income_list_gross(clt_i) = "" then income_list_gross(clt_i) = "0"
		if income_list_net(clt_i) = "" then income_list_net(clt_i) = "0"
		memb_gross_total = Cstr(Cint(memb_gross_total) + Cint(income_list_gross(clt_i)))
		memb_net_total = Cstr(Cint(memb_net_total) + Cint(income_list_net(clt_i)))
		Next

		'Unearn calculations'
		case_memb_unea_total = 0

		f = 0
		For clt_i = 0 to total_clients
		IF all_clients_array(clt_i, 0) <> "" THEN
			IF all_clients_array(clt_i, 1) = 1 THEN unea_ssi_list(f) = ssi_income(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN unea_rsdi_list(f) = rsdi_income(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN unea_child_support_list(f) = child_support(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN unea_other_list(f) = other_unea(clt_i)
			IF all_clients_array(clt_i, 1) = 1 THEN f = f + 1
		End If
		Next
		For clt_i = 0 to total_clients
			If ssi_income(clt_i) = "" then ssi_income(clt_i) = "0"
			If rsdi_income(clt_i) = "" then rsdi_income(clt_i) = "0"
			If child_support(clt_i) = "" then child_support(clt_i) = "0"
			If other_unea(clt_i) = "" then other_unea(clt_i) = "0"
		Next
		FOR clt_i = 0 to 49
		if unea_ssi_list(clt_i) = "" then unea_ssi_list(clt_i) = "0"
		if unea_rsdi_list(clt_i) = "" then unea_rsdi_list(clt_i) = "0"
		if unea_child_support_list(clt_i) = "" then unea_child_support_list(clt_i) = "0"
		if unea_other_list(clt_i) = "" then unea_other_list(clt_i) = "0"
		case_memb_unea_total = Cstr(Cint(case_memb_unea_total) + Cint(unea_ssi_list(clt_i)) + Cint(unea_rsdi_list(clt_i)) + Cint(unea_child_support_list(clt_i)) + Cint(unea_other_list(clt_i)))
		Next






		If fs_prog = true then
			fs_results = "FS: $" & fs_amt_total & "   "
		Else
			fs_results = ""
		End If
		If mf_prog = true then
			mf_results = "MFIP: $" & mf_amt_total & "   "
		Else
			mf_results = ""
		End If
		If mf_fs_prog = true then
			mf_fs_results = "MF-FS: $" & mf_fs_amt_total & "   "
		Else
			mf_fs_results = ""
		End If
		If mf_hg_prog = true then
			mf_hg_results = "MF-HG: $" & mf_hg_amt_total & "   "
		Else
			mf_hg_results = ""
		End If

		'living expense total'



		If flat_living_expense = "Yes" then
			flat_living_expense_amt = "500"
		Else
			flat_living_expense_amt = "0"
		End If

		If flat_trans = "Yes" then
			flat_trans_amt = "113.50"
		Else
			flat_trans_amt = "0"
		End If

		If flat_phone = "Yes" then
			flat_phone_amt = "38"
		Else
			flat_phone_amt = "0"
		End If

		total_expense = Cstr(Cint(shel_paid) + Cint(hest_paid) + Cint(actual_paid) + Cint(other_paid) + Cint(fs_expense) + Cint(flat_living_expense_amt) + Cint(flat_trans_amt) + Cint(flat_phone_amt))

		'%50 test'
		total_gross_income = Cstr(Cint(memb_gross_total))
		total_net_income = Cstr(Cint(memb_net_total))
		unearn_income_total_w_grants = Cstr(Cint(mf_hg_amt_total) + Cint(mf_amt_total) + Cint(case_memb_unea_total))
		total_net_income_for_test = Cstr(Cint(memb_net_total) + Cint(unearn_income_total_w_grants))
		half_total_net_income = Cstr(Cint(total_net_income_for_test)/2)

		'total_cash_grant = Cint(mf_amt_total) + Cint(mf_amt_total) + Cint(mf_amt_total)
		shel_max = Cstr(Cstr(Cint(rent_mo) * 2) + Cint(court_fees) + Cint(late_fees))
		shel_max_allowed = Cstr(Cint(rent_due) + Cint(late_fees) + Cint(court_fees) + Cint(damage_dep))

		'EA Tests'
		'12 months
		If EMER_available_date = "Currently available" then
		   month_test = ":: 12 month test: PASSED!"
		   test_pass1 = true
		Else
		   month_test = ":: 12 month test: FAILED!"
		   test_pass1 = false
		End If
			'FPG test
		If Cint(total_gross_income) <= Cint(FPG_size) then
		   FPG_test = ":: FPG test: PASSED!"
		   test_pass2 = true
		Else
		   FPG_test = ":: FPG test: FAILED! :: Over by $" & Cstr(Cint(total_gross_income) - Cint(FPG_size))
		   test_pass2 = false
		End If
		'50% test
		If Cint(half_total_net_income) <= Cint(total_expense) then
		   percent_test = ":: 50% test: PASSED!"
		   test_pass3 = true
		Else
		   percent_test = ":: 50% test: FAILED! :: short by $" & Cstr(Cint(half_total_net_income) - Cint(total_expense))
		   test_pass3 = false
		End If
		'CostEff test
		If Cint(total_net_income_for_test) >= Cint(rent_mo) then
		   cost_eff_test = ":: Cost-Eff: PASSED!"
		   test_pass4 = true
		Else
		   cost_eff_test = ":: Cost-Eff: FAILED! :: rent over net by $" & Cstr(Cint(rent_mo) - Cint(total_net_income_for_test))
		   test_pass4 = false
		End If
		'Under Shelter Maximum
		If Cint(shel_max) >= Cint(shel_max_allowed) then
		   shel_max_test = ":: Under Shelter Max: PASSED!"
		   test_pass5 = true
		Else
		   shel_max_test = ":: Under Shelter Max: FAILED! :: MAX is: $" & shel_max
		   test_pass5 = false
		End If

		'Under Utilities Maximum
		If Cint(hest_due) <= 1800 then
			hest_due_test = ":: Under Utilities Max: PASSED!"
		   test_pass6 = true
		Else
		   hest_due_test = ":: Under Utilities Max: FAILED! :: over $" & Cstr(Cint(hest_due) - 1800)
		   test_pass6 = false
		End If

        Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 546, 430, "Emergency Screening dialog"
		  DropListBox 50, 55, 65, 45, "Select One"+chr(9)+"Yes"+chr(9)+"No", active_case
		  ButtonGroup ButtonPressed
		    PushButton 10, 105, 40, 15, "HH Memb", MEMB_number
		  CheckBox 15, 140, 40, 10, "Shelter", shelter_check
		  CheckBox 65, 140, 90, 10, "Subsidized(check if yes)", subsidized_check
		  CheckBox 170, 140, 35, 10, "Utility", utility_check
		  ButtonGroup ButtonPressed
		    PushButton 10, 175, 40, 15, "Calc", shelter_button
		    PushButton 10, 225, 40, 15, "Calc", ei_button
		    PushButton 175, 225, 40, 15, "Calc", unearn_income
		    PushButton 10, 320, 40, 15, "Calc", expense_button
		    PushButton 325, 230, 25, 10, "JOBS", JOBS_button
		    PushButton 325, 240, 25, 10, "UNEA", UNEA_button
		    PushButton 255, 15, 30, 10, "prev pnl", prev_panel_button
		    PushButton 285, 15, 30, 10, "next pnl", next_panel_button
		    PushButton 165, 15, 25, 10, "MEMB", MEMB_button
		    PushButton 190, 15, 25, 10, "TYPE", TYPE_button
		    PushButton 215, 15, 25, 10, "PROG", PROG_button
		  GroupBox 320, 210, 35, 45, "Pnls"
		  GroupBox 210, 130, 60, 25, "Locations"
		  GroupBox 250, 5, 70, 25, "STAT-Nav:"
		  GroupBox 160, 5, 85, 25, "other STAT panels:"
		  Text 60, 110, 50, 10, "Size: " & HH_size
		  Text 120, 110, 100, 10, "200% FPG =" & chr(9) & FPG_size
		  GroupBox 5, 95, 235, 30, "House Hold Composition"
		  GroupBox 5, 130, 200, 25, "Emergency Type:"
		  CheckBox 50, 75, 50, 10, "Same Day", same_day
		  Text 10, 60, 40, 10, "Active case:"
		  GroupBox 5, 210, 160, 45, "EARN INCOME"
		  Text 60, 225, 40, 10, "Gross Total:"
		  Text 60, 240, 35, 10, "Net Total:"
		  ButtonGroup ButtonPressed
		    PushButton 240, 140, 25, 10, "ADDR", ADDR_button
		  Text 105, 225, 50, 10, "$" & total_gross_income
		  Text 105, 240, 50, 10, "$" & total_net_income
		  ButtonGroup ButtonPressed
		    PushButton 120, 55, 25, 10, "CURR", CURR_button
		  GroupBox 160, 35, 140, 45, "EA/EGA disbursement 12 months HIST"
		  ButtonGroup ButtonPressed
		    PushButton 305, 65, 25, 10, "INQX", INQX_button
		  GroupBox 115, 100, 120, 20, ""
		  GroupBox 55, 100, 55, 20, ""
		  Text 165, 50, 130, 10, "Last Used: " & EMER_last_used_dates
		  Text 165, 65, 130, 10, "Available: " & EMER_available_date
		  GroupBox 245, 95, 110, 30, "Thrifty Food Plan:"
		  Text 255, 110, 50, 10, "$" & thrifty_food
		  GroupBox 250, 100, 100, 20, ""
		  GroupBox 5, 5, 150, 25, "Period of Evaluation:"
		  Text 10, 15, 65, 10, "From: " & dateadd("d", -30, app_date)
		  Text 80, 15, 70, 10, "To: " & dateadd("d", -1, app_date)
		  GroupBox 170, 210, 145, 45, "UNEARN INCOME"
		  Text 225, 225, 20, 10, "Total:"
		  ButtonGroup ButtonPressed
		    PushButton 325, 220, 25, 10, "BUSI", BUSI_button
		  Text 255, 225, 50, 10, "$" & case_memb_unea_total
		  GroupBox 5, 160, 160, 45, "Shelter Information/Expenses Due"
		  Text 60, 175, 40, 10, "rent/mo:"
		  Text 60, 190, 35, 10, "total due:"
		  ButtonGroup ButtonPressed
		    PushButton 215, 140, 25, 10, "SHEL", SHEL_button
		  Text 105, 175, 50, 10, "$" & rent_mo
		  Text 105, 190, 50, 10, "$" & shel_due
		  EditBox 210, 185, 145, 15, Edit1
		  Text 175, 190, 35, 10, "LandLord:"
		  GroupBox 5, 265, 350, 35, "Programs Issued from EA Eval Period"
		  Text 10, 280, 335, 10, fs_results & mf_results & mf_hg_results & mf_fs_results
		  GroupBox 5, 305, 155, 40, "Living Expenses Paid from EA Eval Period:"
		  Text 60, 325, 90, 10, "$" & total_expense
		  GroupBox 5, 350, 350, 75, ""
		  Text 370, 15, 120, 55, "Gross Earned Income:" & chr(9) & "$" & total_gross_income & vbNewLine & "Net Earned Income:" & chr(9) & "$" & total_net_income & vbNewLine & "Unearned Income:" & chr(9) & "$" & unearn_income_total_w_grants & vbNewLine & "--------------------------------------------------" & vbNewLine & "Total Net Income:" & chr(9) & "$" & total_net_income_for_test & vbNewLine & "50% of Net Income:" & chr(9) & "$" & half_total_net_income & vbNewLine & "EXPENSES TOTAL:" & chr(9) & "$" & total_expense
		  'Elig Determination texts
		  Text 365, 85, 170, 95, month_test & vbNewLine & FPG_test & vbNewLine & percent_test & vbNewLine & cost_eff_test & vbNewLine & shel_max_test & vbNewLine & hest_due_test & vbNewLine & hh_msg
		  If test_pass1 = true and test_pass2 = true and test_pass3 = true and test_pass4 = true and test_pass5 = true and test_pass6 = true and test_pass7 = true then
		      Text 365, 190, 175, 45, "Potential Elig?:  ::YES::"
		  Else
		      Text 365, 190, 175, 45, "Potential Elig?:  ::NO::" & vbNewLine & "Please resolve the 'FAILED!' above tests to be eligible"
		  End If
		  EditBox 365, 375, 175, 15, verification_request
          Text 365, 365, 105, 10, "Additional Verification Request:"
          GroupBox 360, 5, 185, 420, "Eligibility Determination:"
          ButtonGroup ButtonPressed
            OkButton 400, 400, 50, 15
            CancelButton 455, 400, 50, 15
          GroupBox 5, 25, 150, 65, ""
          Text 30, 40, 85, 10, "PROGRAM TYPE:     EA"
          GroupBox 95, 30, 30, 20, ""
          Text 370, 290, 170, 65, Verif_request_list
          Text 370, 220, 170, 60, Verif_received_list
          GroupBox 360, 210, 185, 75, "Verification Received"
          GroupBox 360, 280, 185, 80, "Verification Requested"
          GroupBox 360, 75, 185, 110, "EA TESTS Results"
        EndDialog

        Dialog Dialog1
		cancel_confirmation
		VERIF_BUTTONS
	Loop
End If
