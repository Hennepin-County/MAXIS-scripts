'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARY BLOCK================================================================================================

'These are funcitons that I borrowed from and may need some update.
'display_HEST_information
'access_HEST_panel

Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(original_month, original_year)
future_months_check = checked

'INITIAL Dialog - case number, footer month, worker signature
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 191, 125, "Case Number"
  EditBox 95, 10, 70, 15, MAXIS_case_number
  EditBox 105, 30, 15, 15, original_month
  EditBox 125, 30, 15, 15, original_year
  CheckBox 15, 50, 140, 10, "Check here to have the script update all", future_months_check
  EditBox 10, 85, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 105, 50, 15
    CancelButton 135, 105, 50, 15
  Text 10, 15, 85, 10, "Enter your case number:"
  Text 10, 35, 90, 10, "Starting Footer Month/Year:"
  Text 25, 60, 120, 10, "future months and send through BG."
  Text 10, 75, 65, 10, "Worker Signature:"
EndDialog

'calling the dialog
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_confirmation

        If IsNumeric(MAXIS_case_number) = FALSE or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Enter a valid case number."       'confirming a valid case number
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(original_month, original_year, err_msg, "*")
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature for your case notes."                        'confirming there is a worker signature

        original_month = trim(original_month)       'cleaning up the entry here
        original_year = trim(original_year)
        If len(original_year) <> 2 or len(original_month) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a 2 digit footer month and year."          'forcing 2 digit month and year to be entered
		If err_msg <> "" Then MsgBox "Resolve:" & vbCr & err_msg
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If future_months_check = checked Then call date_array_generator(original_month, original_year, date_array)
If future_months_check = unchecked Then
	footer_date =  original_month & "/1/" & original_year
	footer_date = DateAdd("d", 0, footer_date)
	Dim date_array(0)
	date_array(0)= footer_date
End If

MAXIS_footer_month = original_month     'setting the footer month and year back to what was entered in the dialog.
MAXIS_footer_year = original_year       'this is split out for the option of having seperate handling prior to the reassignment for working in current month if needed

Call back_to_SELF                       'need to gather some detail to have the correct script run

Call navigate_to_MAXIS_screen_review_PRIV("STAT", "HEST", is_this_priv)

EMReadScreen hest_version, 1, 2, 73
If hest_version = "0" Then utilities_paid = "No"
If hest_version = "1" Then
	utilities_paid = "Yes"

	hest_col = 40
	Do
		EMReadScreen pers_paying, 2, 6, hest_col
		If pers_paying <> "__" Then
			all_persons_paying = all_persons_paying & ", " & pers_paying
		Else
			exit do
		End If
		hest_col = hest_col + 3
	Loop until hest_col = 70
	If left(all_persons_paying, 1) = "," Then all_persons_paying = right(all_persons_paying, len(all_persons_paying) - 2)

	EMReadScreen choice_date, 8, 7, 40
	EMReadScreen actual_initial_exp, 8, 8, 61

	EMReadScreen retro_heat_ac_yn, 1, 13, 34
	EMReadScreen retro_heat_ac_units, 2, 13, 42
	EMReadScreen retro_heat_ac_amt, 6, 13, 49
	EMReadScreen retro_electric_yn, 1, 14, 34
	EMReadScreen retro_electric_units, 2, 14, 42
	EMReadScreen retro_electric_amt, 6, 14, 49
	EMReadScreen retro_phone_yn, 1, 15, 34
	EMReadScreen retro_phone_units, 2, 15, 42
	EMReadScreen retro_phone_amt, 6, 15, 49

	EMReadScreen prosp_heat_ac_yn, 1, 13, 60
	EMReadScreen prosp_heat_ac_units, 2, 13, 68
	EMReadScreen prosp_heat_ac_amt, 6, 13, 75
	EMReadScreen prosp_electric_yn, 1, 14, 60
	EMReadScreen prosp_electric_units, 2, 14, 68
	EMReadScreen prosp_electric_amt, 6, 14, 75
	EMReadScreen prosp_phone_yn, 1, 15, 60
	EMReadScreen prosp_phone_units, 2, 15, 68
	EMReadScreen prosp_phone_amt, 6, 15, 75

	choice_date = replace(choice_date, " ", "/")
	If choice_date = "__/__/__" Then choice_date = ""
	actual_initial_exp = trim(actual_initial_exp)
	actual_initial_exp = replace(actual_initial_exp, "_", "")

End If


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 386, 230, "Utilitiy Information"
  DropListBox 190, 20, 50, 45, "Select:"+chr(9)+"Yes"+chr(9)+"No", utilities_paid
  EditBox 125, 40, 125, 15, all_persons_paying
  EditBox 125, 70, 50, 15, choice_date
  EditBox 125, 90, 50, 15, actual_initial_exp
  DropListBox 155, 110, 40, 45, "Select:"+chr(9)+"Yes"+chr(9)+"No", paid_heat
  DropListBox 155, 130, 40, 45, "Select:"+chr(9)+"Yes"+chr(9)+"No", paid_electric
  DropListBox 320, 130, 40, 45, "Select:"+chr(9)+"Yes"+chr(9)+"No", paid_ac
  DropListBox 155, 150, 40, 45, "Select:"+chr(9)+"Yes"+chr(9)+"No", paid_phone
  EditBox 10, 190, 370, 15, notes_on_hest
  ButtonGroup ButtonPressed
    OkButton 275, 210, 50, 15
    CancelButton 330, 210, 50, 15
  GroupBox 10, 10, 370, 160, "Utility Information (SUA - Standard Utility Allowance)"
  Text 15, 25, 175, 10, "Is the Household Responsible to Pay any Utilities?"
  Text 15, 45, 110, 10, "Who is responsible to pay utilities?"
  Text 125, 55, 250, 10, "Enter member references numbers separated by commas who pay utilities."
  Text 70, 75, 55, 10, "FS Choice Date:"
  Text 180, 75, 150, 10, "Date utility payment was determined."
  Text 15, 95, 110, 10, "Actual Expense In Initial Month: $ "
  Text 20, 115, 90, 10, "Does the resident pay for:"
  Text 130, 115, 25, 10, "Heat:"
  Text 125, 135, 30, 10, "Electric:"
  Text 200, 135, 120, 10, "Is Air Conditioning part of the Electric:"
  Text 125, 155, 30, 10, "  Phone:"
  Text 10, 180, 75, 10, "Additional Notes:"
EndDialog

Do
	err_msg = ""

	dialog Dialog1
	cancel_confirmation

	all_persons_paying = trim(all_persons_paying)

	If utilities_paid = "Select:" Then err_msg = err_msg & vbCr & "* Indicate if the household is responsible to pay utilities."
	If utilities_paid = "Yes" Then
		If all_persons_paying = "" Then err_msg = err_msg & vbCr & "* Enter at least 1 reference number for the household member is responsible to pay the utilities."
		If IsDate(choice_date) = False Then err_msg = err_msg & vbCr & "* FS Choice Date must be a valid date."
		If paid_heat <> "Yes" and paid_electric <> "Yes" and paid_ac <> "Yes" and paid_phone <> "Yes" Then err_msg = err_msg & vbCr & "* If utilities are paid, at least one utility type must be yes."
		If paid_ac = "Yes" and paid_electric <> "Yes" Then err_msg = err_msg & vbCr & "* If AC is paid, the Electric must also be paid."
		If paid_electric = "Yes" and paid_ac = "Select:" Then err_msg = err_msg & vbCr & "* Since electric is 'yes', the question about AC must be answered."
	End If
	If err_msg <> "" Then MsgBox "Resolve:" & vbCr & err_msg
Loop until err_msg = ""

retro_heat_ac_yn = " "
prosp_heat_ac_yn = " "
retro_electric_yn = " "
prosp_electric_yn = " "
retro_phone_yn = " "
prosp_phone_yn = " "
If paid_heat = "Yes" or paid_ac = "Yes" Then
	retro_heat_ac_yn = "Y"
	prosp_heat_ac_yn = "Y"
Else
	If paid_electric = "Yes" Then
		retro_electric_yn = "Y"
		prosp_electric_yn = "Y"
	End If
	If paid_phone = "Yes" Then
		retro_phone_yn = "Y"
		prosp_phone_yn = "Y"
	End If
End If
prosp_heat_ac_amt = 0
prosp_electric_amt = 0
prosp_phone_amt = 0

Call back_to_SELF
Call MAXIS_background_check

For each footer_month in date_array

	MAXIS_footer_month = datepart("m", footer_month) 'Need to assign footer month / year each time through
	If len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
	MAXIS_footer_year = right(datepart("YYYY", footer_month), 2)

	EMReadScreen hest_check, 4, 2, 53
	If hest_check <> "HEST" Then
		Call MAXIS_background_check
		Call navigate_to_MAXIS_screen("STAT", "HEST")
	End If

	EMReadScreen hest_version, 1, 2, 73

	If utilities_paid = "No" Then
		If hest_version = "1" Then
			EMWriteScreen "DEL", 20, 71
			PF9
			transmit
		End If
	End If
	If utilities_paid = "Yes" Then

		If hest_version = "1" Then PF9
		If hest_version = "0" Then
			EMWriteScreen "NN", 20, 79
			transmit
		End If

		all_persons_paying = trim(all_persons_paying)
		If all_persons_paying <> "" Then
			If InStr(all_persons_paying, ",") = 0 Then
				persons_array = array(all_persons_paying)
			Else
				persons_array = split(all_persons_paying, ",")
			End If

			For Hest_col = 40 to 67 Step 3
				EMWriteScreen "  ", 6, hest_col
			Next

			hest_col = 40
			for each pers_paying in persons_array
				EMWriteScreen trim(pers_paying), 6, hest_col
				hest_col = hest_col + 3
			Next

			For row = 13 to 15
				EMWriteScreen "  ", row, 42
				EMWriteScreen "  ", row, 68
			Next

			If IsDate(choice_date) = True Then Call create_mainframe_friendly_date(choice_date, 7, 40, "YY")
			Call clear_line_of_text(8, 61)
			EMWriteScreen actual_initial_exp, 8, 61

			EMWriteScreen retro_heat_ac_yn, 13, 34
			If retro_heat_ac_yn = "Y" Then EMWriteScreen "01", 13, 42
			EMWriteScreen retro_electric_yn, 14, 34
			If retro_electric_yn = "Y" Then EMWriteScreen "01", 14, 42
			EMWriteScreen retro_phone_yn, 15, 34
			If retro_phone_yn = "Y" Then EMWriteScreen "01", 15, 42

			EMWriteScreen prosp_heat_ac_yn, 13, 60
			If prosp_heat_ac_yn = "Y" Then EMWriteScreen "01", 13, 68
			EMWriteScreen prosp_electric_yn, 14, 60
			If prosp_electric_yn = "Y" Then EMWriteScreen "01", 14, 68
			EMWriteScreen prosp_phone_yn, 15, 60
			If prosp_phone_yn = "Y" Then EMWriteScreen "01", 15, 68

			transmit

			EMReadScreen retro_heat_ac_amt, 6, 13, 49
			EMReadScreen retro_electric_amt, 6, 14, 49
			EMReadScreen retro_phone_amt, 6, 15, 49

			EMReadScreen prosp_heat_ac_amt, 6, 13, 75
			EMReadScreen prosp_electric_amt, 6, 14, 75
			EMReadScreen prosp_phone_amt, 6, 15, 75

			retro_heat_ac_amt = trim(retro_heat_ac_amt)
			If retro_heat_ac_amt = "" Then retro_heat_ac_amt = 0
			retro_heat_ac_amt = retro_heat_ac_amt * 1
			retro_electric_amt = trim(retro_electric_amt)
			If retro_electric_amt = "" Then retro_electric_amt = 0
			retro_electric_amt = retro_electric_amt * 1
			retro_phone_amt = trim(retro_phone_amt)
			If retro_phone_amt = "" Then retro_phone_amt = 0
			retro_phone_amt = retro_phone_amt * 1

			prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
			If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
			prosp_heat_ac_amt = prosp_heat_ac_amt * 1
			prosp_electric_amt = trim(prosp_electric_amt)
			If prosp_electric_amt = "" Then prosp_electric_amt = 0
			prosp_electric_amt = prosp_electric_amt * 1
			prosp_phone_amt = trim(prosp_phone_amt)
			If prosp_phone_amt = "" Then prosp_phone_amt = 0
			prosp_phone_amt = prosp_phone_amt * 1
		End If
	End If
	If future_months_check = unchecked Then Exit For
	PF3
	EMReadScreen wrap_check, 4, 2, 46
	If wrap_check = "WRAP" Then
		Call write_value_and_transmit("Y", 16, 54)
		EMReadScreen pnlp_check, 4, 2, 53
		If pnlp_check = "PNLP" Then
			Call write_value_and_transmit("HEST", 20, 71)
		Else
			Call write_value_and_transmit("N", 16, 54)
		End If
	End If
Next

Call back_to_SELF

member_note_detail = replace(all_persons_paying, " ", "")
If right(member_note_detail, 1) = "," Then member_note_detail = left(member_note_detail, len(member_note_detail)-1)
member_note_detail = replace(member_note_detail, ",", ", MEMB ")
member_note_detail = "MEMB " & member_note_detail

total_hest = 0
If prosp_heat_ac_amt = 0 Then
	total_hest = prosp_electric_amt + prosp_phone_amt
Else
	total_hest = prosp_heat_ac_amt
End If


Call start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE("Utility Payment Details")
If utilities_paid = "No" Then
	Call write_variable_in_CASE_NOTE("Resident indicated they are not responsible to pay any utilities.")
	Call write_variable_in_CASE_NOTE("HEST panel deleted.")
End If
If utilities_paid = "Yes" Then
	Call write_bullet_and_variable_in_CASE_NOTE("Responsible Members", member_note_detail)
	Call write_bullet_and_variable_in_CASE_NOTE("Selection Date", choice_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Initial Month Expense", actual_initial_exp)
	Call write_variable_in_CASE_NOTE("* Utilities Paid:")
	If paid_heat = "Yes" Then Call write_variable_in_CASE_NOTE("  - Heat")
	If paid_electric = "Yes" Then
		If paid_ac = "Yes" Then Call write_variable_in_CASE_NOTE("  - Electric - Includes A/C at some point in the year")
		If paid_ac = "No" Then Call write_variable_in_CASE_NOTE("  - Electric - A/C is not paid with the electric bill")
	End If
	If paid_phone = "Yes" Then Call write_variable_in_CASE_NOTE("  - Phone")
	Call write_variable_in_CASE_NOTE("* HEST Updated:")
	Call write_bullet_and_variable_in_CASE_NOTE("Total Utilities Expense", "$ " & total_hest)
	If prosp_heat_ac_yn = "Y" Then
		Call write_variable_in_CASE_NOTE("  - Heat/AC Standard: $ " & prosp_heat_ac_amt)
	Else
		If prosp_electric_yn = "Y" Then Call write_variable_in_CASE_NOTE("  - Electric: $ " & prosp_electric_amt)
		If prosp_phone_yn = "Y" Then Call write_variable_in_CASE_NOTE("  - Phone: $ " & prosp_phone_amt)
	End If
End If
Call write_bullet_and_variable_in_CASE_NOTE("Notes", notes_on_hest)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

Call script_end_procedure("")