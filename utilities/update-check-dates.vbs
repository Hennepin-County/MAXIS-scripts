'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - Update Check Dates.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 180                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("03/02/2022", "BUG FIX - Semi Monthly income sources were not correctly determining the pay dates and were not updating panels correctly.", "Casey Love, Hennepin County")
call changelog_update("08/07/2020", "Bug Fix: Sometimes there was an error after selecting which income needs to be updated. Updated the script to not reach the error.", "Casey Love, Hennepin County")
call changelog_update("05/19/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT ====================================================================================================================
EMConnect ""
Call check_for_MAXIS(true)

'Autofilling information
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Call back_to_SELF                       'go back to self to read the region.
developer_mode = FALSE                  'defaulting developer to false because we are usually not in developer mode
EMReadScreen MX_region, 12, 22, 48      'reading for the region
MX_region = trim(MX_region)             'formatting the region
If MX_region = "INQUIRY DB" Then        'This is what INQUIRY looks like on SELF.
    'We are going to confirm HERE that the worker meant to run this in inquiry If not, the script run will end.
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Panels cannot be updated." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
    developer_mode = TRUE               'If thes cript didn't end and we were in inquiry, we are automatically in developer mode
End If
' If worker_signature = "UUDDLRLRBA" Then developer_mode = TRUE           'Use of the konami code in worker_signature will also cause the script to run in developer mode
If developer_mode = TRUE then MsgBox "Developer Mode ACTIVATED!"        'developer mode prevents actions from being taken in MAXIS
If developer_mode = TRUE Then script_run_lowdown = "Run in INQUIRY" & vbCr & vbCr     'adding this to any error reporting


Do
    Do
        err_msg = ""

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 276, 120, "Case Number and Footer Month"
          EditBox 120, 35, 60, 15, MAXIS_case_number
          EditBox 120, 55, 20, 15, MAXIS_footer_month
          EditBox 150, 55, 20, 15, MAXIS_footer_year
          ButtonGroup ButtonPressed
            OkButton 220, 80, 50, 15
            CancelButton 220, 100, 50, 15
          Text 5, 10, 265, 20, "This script will update the pay dates only for future JOBS and UNEA panels as MAXIS requires paydates to match the footer month to generate eligibility results."
          Text 10, 40, 105, 10, "MAXIS Case Number to Update:"
          Text 10, 60, 105, 10, "First month to Update (MM/YY):"
          Text 10, 80, 180, 20, "The script will find the JOBS and UNEA panels after entering the case number and footer month and year. "
          Text 10, 105, 170, 10, "Update will happen through Current Month plus 1."
          Text 145, 60, 5, 10, "/"
        EndDialog

        dialog Dialog1
        cancel_without_confirmation

        call validate_MAXIS_case_number(err_msg, "*")

        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call MAXIS_footer_month_confirmation
call MAXIS_background_check

const panel_type_const      = 0
const panel_member_const    = 1
const panel_instance_const  = 2
const panel_freq_const      = 3
const panel_weekday_const   = 4
const panel_pay_amt_const   = 5
const panel_update_checkbox = 6
const panel_known_paydate   = 7
const panel_name_const      = 8
const semi_mo_pay_one       = 9
const semi_mo_pay_two       = 10
const check_hours_const     = 11
const panel_notes_const     = 12

Dim PANELS_ARRAY()
ReDim PANELS_ARRAY(panel_notes_const, 0)

day_list = "?"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"13"+chr(9)+"14"+chr(9)+"15"+chr(9)+"16"+chr(9)+"17"+chr(9)+"18"+chr(9)+"19"+chr(9)+"20"+chr(9)+"21"+chr(9)+"22"+chr(9)+"23"+chr(9)+"24"+chr(9)+"25"+chr(9)+"26"+chr(9)+"27"+chr(9)+"28"+chr(9)+"29"+chr(9)+"30"
second_day_list = day_list+chr(9)+"Last"

Call navigate_to_MAXIS_screen("STAT", "SUMM")
EMWriteScreen "PNLI", 20, 71
transmit

stat_row = 2
panel_counter = 0
instance_counter = 1
Do
    command_to_go_to = ""
    EMReadScreen panel_name, 4, stat_row, 5
    If panel_name = "JOBS" or panel_name = "UNEA" Then
        ReDim Preserve PANELS_ARRAY(panel_notes_const, panel_counter)
        EMReadScreen member_number, 2, stat_row, 10
        If prev_panel_name <> panel_name OR prev_panel_member <> member_number Then instance_counter = 1

        PANELS_ARRAY(panel_type_const, panel_counter) = panel_name
        PANELS_ARRAY(panel_member_const, panel_counter) = member_number
        PANELS_ARRAY(panel_instance_const, panel_counter) = "0" & instance_counter
        instance_counter = instance_counter + 1

        prev_panel_name = panel_name
        prev_panel_member = member_number
        panel_counter = panel_counter + 1
    End If

    stat_row = stat_row + 1
    If stat_row >= 20 Then
        EMReadScreen command_to_go_to, 4, 20, 71
        transmit
        stat_row = 2
    End If
Loop until command_to_go_to = "PNLE"

For view_panel = 0 to UBound(PANELS_ARRAY, 2)
    EMWriteScreen PANELS_ARRAY(panel_type_const, view_panel), 20, 71
    EMWriteScreen PANELS_ARRAY(panel_member_const, view_panel), 20, 76
    EMWriteScreen PANELS_ARRAY(panel_instance_const, view_panel), 20, 79
    transmit

    If PANELS_ARRAY(panel_type_const, view_panel) = "JOBS" Then
        EMReadScreen employer_name, 30, 7, 42
        PANELS_ARRAY(panel_name_const, view_panel) = replace(employer_name, "_", "")
        'look for frequency in SNAP PIC
        EMWriteScreen "X", 19, 38
        transmit

        EMReadScreen snap_pic_pay_freq, 1, 5, 64
        If snap_pic_pay_freq <> "_" Then PANELS_ARRAY(panel_freq_const, view_panel) = snap_pic_pay_freq

        EMReadScreen pic_paydate, 8, 9, 13
        If pic_paydate <> "__ __ __" Then PANELS_ARRAY(panel_known_paydate, view_panel) = replace(pic_paydate, " ", "/")
        PF3

        'Look for frequency in HC popup
        EMReadScreen jobs_pay_freq, 1, 18, 35
        If jobs_pay_freq <> "_" Then
            PANELS_ARRAY(panel_freq_const, view_panel) = jobs_pay_freq
        End If

        'Use the number of checks to determine frequency
        EMReadScreen paycheck_one, 8, 12, 54
        If PANELS_ARRAY(panel_known_paydate, view_panel) = "" AND paycheck_one <> "__ __ __" Then
            paycheck_one = replace(paycheck_one, " ", "/")
            check_month = DatePart("m", paycheck_one)
            check_month = right("00"&check_month, 2)
            If check_month = MAXIS_footer_month Then PANELS_ARRAY(panel_known_paydate, view_panel) = paycheck_one
        End If

        If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE Then PANELS_ARRAY(panel_weekday_const, view_panel) = WeekDayName(WeekDay(PANELS_ARRAY(panel_known_paydate, view_panel)))

        'Finding the amount
        EMReadScreen amount_one,   8, 12, 67
        EMReadScreen amount_two,   8, 13, 67
        EMReadScreen amount_three, 8, 14, 67
        EMReadScreen amount_four,  8, 15, 67
        EMReadScreen amount_five,  8, 16, 67
        EMReadScreen prosp_hours,  3, 18, 72
        If prosp_hours = "___" then prosp_hours = 0
        If prosp_hours = "?__" then prosp_hours = 0
        prosp_hours = prosp_hours * 1

        If amount_five <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_five
            PANELS_ARRAY(check_hours_const, view_panel) = prosp_hours/5
        ElseIf amount_four <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_four
            PANELS_ARRAY(check_hours_const, view_panel) = prosp_hours/4
        ElseIf amount_three <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_three
            PANELS_ARRAY(check_hours_const, view_panel) = prosp_hours/3
        ElseIf amount_two <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_two
            PANELS_ARRAY(check_hours_const, view_panel) = prosp_hours/2
        ElseIf amount_one <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_one
            PANELS_ARRAY(check_hours_const, view_panel) = prosp_hours
        End If

    End If

    If PANELS_ARRAY(panel_type_const, view_panel) = "UNEA" Then
        EMReadScreen panel_income, 2, 5, 37
        If panel_income = "01" Then PANELS_ARRAY(panel_name_const, view_panel) = "01 - RSDI, Disa"
        If panel_income = "02" Then PANELS_ARRAY(panel_name_const, view_panel) = "02 - RSDI, No Disa"
        If panel_income = "03" Then PANELS_ARRAY(panel_name_const, view_panel) = "03 - SSI"
        If panel_income = "06" Then PANELS_ARRAY(panel_name_const, view_panel) = "06 - Non-MN PA"
        If panel_income = "11" Then PANELS_ARRAY(panel_name_const, view_panel) = "11 - VA Disability Benefit"
        If panel_income = "12" Then PANELS_ARRAY(panel_name_const, view_panel) = "12 - VA Pension"
        If panel_income = "13" Then PANELS_ARRAY(panel_name_const, view_panel) = "13 - VA Other"
        If panel_income = "38" Then PANELS_ARRAY(panel_name_const, view_panel) = "38 - VA Aid & Attendance"
        If panel_income = "14" Then PANELS_ARRAY(panel_name_const, view_panel) = "14 - Unemployment Insurance"
        If panel_income = "15" Then PANELS_ARRAY(panel_name_const, view_panel) = "15 - Worker's Comp"
        If panel_income = "16" Then PANELS_ARRAY(panel_name_const, view_panel) = "16 - Railroad Retirement"
        If panel_income = "17" Then PANELS_ARRAY(panel_name_const, view_panel) = "17 - Other Retirement"
        If panel_income = "18" Then PANELS_ARRAY(panel_name_const, view_panel) = "18 - Military Allotment"
        If panel_income = "19" Then PANELS_ARRAY(panel_name_const, view_panel) = "19 - FC Child Requesting FS"
        If panel_income = "20" Then PANELS_ARRAY(panel_name_const, view_panel) = "20 - FC Child Not Req FS"
        If panel_income = "21" Then PANELS_ARRAY(panel_name_const, view_panel) = "21 - FC Adult Requesting FS"
        If panel_income = "22" Then PANELS_ARRAY(panel_name_const, view_panel) = "22 - FC Adult Not Req FS"
        If panel_income = "23" Then PANELS_ARRAY(panel_name_const, view_panel) = "23 - Dividends"
        If panel_income = "24" Then PANELS_ARRAY(panel_name_const, view_panel) = "24 - Interest"
        If panel_income = "25" Then PANELS_ARRAY(panel_name_const, view_panel) = "25 - Cnt Gifts Or Prizes"
        If panel_income = "26" Then PANELS_ARRAY(panel_name_const, view_panel) = "26 - Strike Benefit"
        If panel_income = "27" Then PANELS_ARRAY(panel_name_const, view_panel) = "27 - Contract for Deed"
        If panel_income = "28" Then PANELS_ARRAY(panel_name_const, view_panel) = "28 - Illegal Income"
        If panel_income = "29" Then PANELS_ARRAY(panel_name_const, view_panel) = "29 - Other Countable"
        If panel_income = "30" Then PANELS_ARRAY(panel_name_const, view_panel) = "30 - Infrequent <30 Not Counted"
        If panel_income = "31" Then PANELS_ARRAY(panel_name_const, view_panel) = "31 - Other FS Only"
        If panel_income = "08" Then PANELS_ARRAY(panel_name_const, view_panel) = "08 - Direct Child Support"
        If panel_income = "35" Then PANELS_ARRAY(panel_name_const, view_panel) = "35 - Direct Spousal Support"
        If panel_income = "36" Then PANELS_ARRAY(panel_name_const, view_panel) = "36 - Disbursed Child Support"
        If panel_income = "37" Then PANELS_ARRAY(panel_name_const, view_panel) = "37 - Disbursed Spousal Sup"
        If panel_income = "39" Then PANELS_ARRAY(panel_name_const, view_panel) = "39 - Disbursed CS Arrears"
        If panel_income = "40" Then PANELS_ARRAY(panel_name_const, view_panel) = "40 - Disbursed Spsl Sup Arrears"
        If panel_income = "43" Then PANELS_ARRAY(panel_name_const, view_panel) = "43 - Disbursed Excess CS"
        If panel_income = "44" Then PANELS_ARRAY(panel_name_const, view_panel) = "44 - MSA - Excess Inc for SSI"
        If panel_income = "45" Then PANELS_ARRAY(panel_name_const, view_panel) = "45 - County 88 Child Support"
        If panel_income = "46" Then PANELS_ARRAY(panel_name_const, view_panel) = "46 - County 88 Gaming"
        If panel_income = "47" Then PANELS_ARRAY(panel_name_const, view_panel) = "47 - Counted Tribal Income"
        If panel_income = "48" Then PANELS_ARRAY(panel_name_const, view_panel) = "48 - Trust Income"
        If panel_income = "49" Then PANELS_ARRAY(panel_name_const, view_panel) = "49 - Non-Recurring Income > $60 per quarter"

        'look for frequency in SNAP PIC
        EMWriteScreen "X", 10, 26
        transmit

        EMReadScreen snap_pic_pay_freq, 1, 5, 64
        If snap_pic_pay_freq <> "_" Then PANELS_ARRAY(panel_freq_const, view_panel) = snap_pic_pay_freq

        EMReadScreen pic_paydate, 8, 9, 13
        If pic_paydate <> "__ __ __" Then PANELS_ARRAY(panel_known_paydate, view_panel) = replace(pic_paydate, " ", "/")
        PF3

        'Look for frequency in HC popup
        If PANELS_ARRAY(panel_freq_const, view_panel) = "" Then
            EMWriteScreen "X", 6, 56
            transmit

            EMReadScreen hc_pic_pay_freq, 1, 10, 63
            If hc_pic_pay_freq <> "_" Then
                PANELS_ARRAY(panel_freq_const, view_panel) = hc_pic_pay_freq
            End If
            PF3
        End If

        'Use the number of checks to determine frequency
        EMReadScreen paycheck_one,   8, 13, 54
        EMReadScreen paycheck_two,   8, 14, 54
        EMReadScreen paycheck_three, 8, 15, 54
        EMReadScreen paycheck_four,  8, 16, 54
        EMReadScreen paycheck_five,  8, 17, 54
        If PANELS_ARRAY(panel_freq_const, view_panel) = "" Then
            If paycheck_five <> "__ __ __" OR paycheck_four <> "__ __ __" Then
                PANELS_ARRAY(panel_freq_const, view_panel) = "4"
            ElseIf paycheck_three <> "__ __ __" Then
                PANELS_ARRAY(panel_freq_const, view_panel) = "3"
            ElseIf paycheck_two <> "__ __ __" Then
                PANELS_ARRAY(panel_freq_const, view_panel) = "2"
            ElseIf paycheck_one <> "__ __ __" Then
                PANELS_ARRAY(panel_freq_const, view_panel) = "1"
            End If
        End If
        If PANELS_ARRAY(panel_known_paydate, view_panel) = "" AND paycheck_one <> "__ __ __" Then
            paycheck_one = replace(paycheck_one, " ", "/")
            check_month = DatePart("m", paycheck_one)
            check_month = right("00"&check_month, 2)
            If check_month = MAXIS_footer_month Then PANELS_ARRAY(panel_known_paydate, view_panel) = paycheck_one
        End If

        If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE Then PANELS_ARRAY(panel_weekday_const, view_panel) = WeekDayName(WeekDay(PANELS_ARRAY(panel_known_paydate, view_panel)))

        'Finding the amount
        EMReadScreen amount_one,   8, 13, 68
        EMReadScreen amount_two,   8, 14, 68
        EMReadScreen amount_three, 8, 15, 68
        EMReadScreen amount_four,  8, 16, 68
        EMReadScreen amount_five,  8, 17, 68

        If amount_five <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_five
        ElseIf amount_four <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_four
        ElseIf amount_three <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_three
        ElseIf amount_two <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_two
        ElseIf amount_one <> "________" Then
            PANELS_ARRAY(panel_pay_amt_const, view_panel) = amount_one
        End If

    End If

    If PANELS_ARRAY(panel_freq_const, view_panel) = "1" Then
        PANELS_ARRAY(panel_freq_const, view_panel) = "1 - Monthly"
        PANELS_ARRAY(panel_weekday_const, view_panel) = "Inconsistent"
        If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE Then
            PANELS_ARRAY(semi_mo_pay_one, view_panel) = DatePart("d", PANELS_ARRAY(panel_known_paydate, view_panel))
            PANELS_ARRAY(panel_known_paydate, view_panel) = ""
        End If
    End If
    If PANELS_ARRAY(panel_freq_const, view_panel) = "2" Then
        PANELS_ARRAY(panel_freq_const, view_panel) = "2 - Semi-Monthly"
        PANELS_ARRAY(panel_weekday_const, view_panel) = "Inconsistent"
        If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE Then
            the_day = DatePart("d", PANELS_ARRAY(panel_known_paydate, view_panel))
			' MsgBox the_day																									'Taking the functionality out to try to guess the dates because it messes up the dialogs.
            ' If the_day < 15 Then PANELS_ARRAY(semi_mo_pay_one, view_panel) = the_day
            ' If the_day > 14 AND the_day < 28 Then PANELS_ARRAY(semi_mo_pay_two, view_panel) = the_day
            ' If the_day > 27 Then PANELS_ARRAY(semi_mo_pay_two, view_panel) = "Last"
			' MsgBox "1 - " & PANELS_ARRAY(semi_mo_pay_one, view_panel) & vbCr & "2 - " & PANELS_ARRAY(semi_mo_pay_two, view_panel)
            PANELS_ARRAY(panel_known_paydate, view_panel) = ""
        End If
    End If
    If PANELS_ARRAY(panel_freq_const, view_panel) = "3" Then PANELS_ARRAY(panel_freq_const, view_panel) = "3 - Biweekly"
    If PANELS_ARRAY(panel_freq_const, view_panel) = "4" Then PANELS_ARRAY(panel_freq_const, view_panel) = "4 - Weekly"
    If PANELS_ARRAY(panel_freq_const, view_panel) = "" Then
        PANELS_ARRAY(panel_freq_const, view_panel) = "Select One..."
        PANELS_ARRAY(panel_weekday_const, view_panel) = "Select One..."
    End If

Next

Do
    Do
        err_msg = ""
        dlg_len = 100 + UBound(PANELS_ARRAY, 2)*20
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 536, dlg_len, "Select Panels to Update"
          Text 10, 10, 425, 10, "Check all of the panels you would like the script to update paydates for:"
          Text 20, 30, 50, 10, "Panel"
          Text 185, 30, 40, 10, "Frequency"
          Text 255, 30, 60, 10, "Known Pay Date"
          Text 320, 30, 60, 10, "Weekday of Pay"
          Text 395, 30, 75, 10, "Always paid on"
          Text 485, 30, 50, 10, "Pay Amount"
          y_pos = 45
          For view_panel = 0 to UBound(PANELS_ARRAY, 2)
              CheckBox 10, y_pos+5, 50, 10, PANELS_ARRAY(panel_type_const, view_panel) & " " & PANELS_ARRAY(panel_member_const, view_panel) & " " & PANELS_ARRAY(panel_instance_const, view_panel), PANELS_ARRAY(panel_update_checkbox, view_panel)
              Text 65, y_pos+5, 105, 10, PANELS_ARRAY(panel_name_const, view_panel)
              DropListBox 185, y_pos, 60, 45, "Select One..."+chr(9)+"1 - Monthly"+chr(9)+"2 - Semi-Monthly"+chr(9)+"3 - Biweekly"+chr(9)+"4 - Weekly", PANELS_ARRAY(panel_freq_const, view_panel)
              EditBox 255, y_pos, 55, 15, PANELS_ARRAY(panel_known_paydate, view_panel)
              DropListBox 320, y_pos, 65, 45, "Select One..."+chr(9)+"Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday"+chr(9)+"Inconsistent", PANELS_ARRAY(panel_weekday_const, view_panel)
              Text 485, y_pos+5, 45, 10, PANELS_ARRAY(panel_pay_amt_const, view_panel)
              DropListBox 395, y_pos, 25, 45, day_list, PANELS_ARRAY(semi_mo_pay_one, view_panel)
              Text 425, y_pos+5, 15, 10, "and"
              DropListBox 440, y_pos, 30, 45, second_day_list, PANELS_ARRAY(semi_mo_pay_two, view_panel)
              y_pos = y_pos + 20
          Next
          ButtonGroup ButtonPressed
            OkButton 425, y_pos+15, 50, 15
            CancelButton 480, y_pos+15, 50, 15
          Text 20, y_pos+5, 385, 25, "This script is only to update the pay dates to match the footer month using the accurate pay dates. It does not update the pay amount or verifciation codes. This is used solely when eligibility results are inhibited but no change has happened to the income budgeting."
        EndDialog

        dialog Dialog1
        cancel_confirmation

        For view_panel = 0 to UBound(PANELS_ARRAY, 2)
            If PANELS_ARRAY(panel_update_checkbox, view_panel) = checked Then
                If PANELS_ARRAY(panel_freq_const, view_panel) = "1 - Monthly" Then
                    If PANELS_ARRAY(semi_mo_pay_one, view_panel) <> "?" AND PANELS_ARRAY(semi_mo_pay_two, view_panel) <> "?" Then err_msg = err_msg & vbNewLine & "* You have indicated two different days of the month that the pay is ALWAYS received but for income received MONTHLY, there can only be one day. Please select just one day."
                    If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE Then
                        If PANELS_ARRAY(semi_mo_pay_one, view_panel) = "?" Then
                            PANELS_ARRAY(semi_mo_pay_one, view_panel) = DatePart("d", PANELS_ARRAY(panel_known_paydate, view_panel))
                        Else
                            If DatePart("d", PANELS_ARRAY(panel_known_paydate, view_panel)) <> PANELS_ARRAY(semi_mo_pay_one, view_panel) Then err_msg = err_msg & vbNewLine & "* For income received MONTHLY the known paydate listed and the day ALWAYS paid on do not match."
                        End If
                    ElseIf PANELS_ARRAY(semi_mo_pay_one, view_panel) = "?" Then
                        err_msg = err_msg & vbNewLine & "* For income received MONTHLY, you must enter either a known paydate or indicate a day in the month pay is ALWAYS received."
                    End If
                End If


                If PANELS_ARRAY(panel_freq_const, view_panel) = "2 - Semi-Monthly" Then
                    If PANELS_ARRAY(semi_mo_pay_one, view_panel) = "?" OR PANELS_ARRAY(semi_mo_pay_two, view_panel) = "?" Then
                        err_msg = err_msg & vbNewLine & "* Since the pay frequency is SEMI-MONTHLY the days the pay is issued in each month need to be indicated in the column 'Always paid on'."
                    Else
                        If PANELS_ARRAY(semi_mo_pay_two, view_panel) <> "Last" Then
							If IsNumeric(PANELS_ARRAY(semi_mo_pay_one, view_panel)) = true then PANELS_ARRAY(semi_mo_pay_one, view_panel) = PANELS_ARRAY(semi_mo_pay_one, view_panel) * 1		'need to ensure the script is reading and comparing these as numbers
							If IsNumeric(PANELS_ARRAY(semi_mo_pay_two, view_panel)) = true then PANELS_ARRAY(semi_mo_pay_two, view_panel) = PANELS_ARRAY(semi_mo_pay_two, view_panel) * 1
							If PANELS_ARRAY(semi_mo_pay_two, view_panel) <= PANELS_ARRAY(semi_mo_pay_one, view_panel) Then
                                number_one = PANELS_ARRAY(semi_mo_pay_one, view_panel)
                                number_two = PANELS_ARRAY(semi_mo_pay_two, view_panel)
                                PANELS_ARRAY(semi_mo_pay_one, view_panel) = number_two
                                PANELS_ARRAY(semi_mo_pay_two, view_panel) = number_one
                                err_msg = err_msg & vbNewLine & "* The days of pay for SEMI-MONTHLY pay frequecny need to be in order. The script has switched:" & vbNewLine & "Semi-Monthly pay day One entered: " & number_one & vbNewLine & "       Changed to: " & PANELS_ARRAY(semi_mo_pay_one, view_panel) & vbNewLine & vbNewLine &_
                                                                "Semi-Monthly pay day Two entered: " & number_two & vbNewLine & "       Changed to: " & PANELS_ARRAY(semi_mo_pay_two, view_panel)
                            End If
                        End If
                    End If
                End If
                If PANELS_ARRAY(panel_freq_const, view_panel) = "3 - Biweekly" Then
                    If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = FALSE Then
                        err_msg = err_msg & vbNewLine & "* For income received BIWEEKLY, a known paydate must be entered."
                    Else
                        If PANELS_ARRAY(panel_weekday_const, view_panel) <> "Select One..." Then
                            If WeekDayName(WeekDay(PANELS_ARRAY(panel_known_paydate, view_panel))) <> PANELS_ARRAY(panel_weekday_const, view_panel) Then err_msg = err_msg & vbNewLine & "* The known pay date you entered does not match up with the weekday provided."
                        End If
                    End If
                End If
                If PANELS_ARRAY(panel_freq_const, view_panel) = "4 - Weekly" Then
                    If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = FALSE Then
                        If PANELS_ARRAY(panel_weekday_const, view_panel) = "Inconsistent" OR PANELS_ARRAY(panel_weekday_const, view_panel) = "Select One..." Then err_msg = err_msg & vbNewLine & "* For an income source that is received WEEKLY you must either provide a known pay date or the day of the week the pay is received on."
                    Else
                        If WeekDayName(WeekDay(PANELS_ARRAY(panel_known_paydate, view_panel))) <> PANELS_ARRAY(panel_weekday_const, view_panel) Then err_msg = err_msg & vbNewLine & "* The known pay date you entered does not match up with the weekday provided."
                    End If
                End If
            End If
        Next
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

For view_panel = 0 to UBound(PANELS_ARRAY, 2)
	If PANELS_ARRAY(panel_update_checkbox, view_panel) = checked Then
	    If PANELS_ARRAY(panel_freq_const, view_panel) = "1 - Monthly" Then
	        PANELS_ARRAY(panel_known_paydate, view_panel) = MAXIS_footer_month & "/" & PANELS_ARRAY(semi_mo_pay_one, view_panel) & "/" & MAXIS_footer_year
	    End If
	    If PANELS_ARRAY(panel_freq_const, view_panel) = "2 - Semi-Monthly" Then
	        PANELS_ARRAY(panel_known_paydate, view_panel) = MAXIS_footer_month & "/" & PANELS_ARRAY(semi_mo_pay_one, view_panel) & "/" & MAXIS_footer_year
			If IsNumeric(PANELS_ARRAY(semi_mo_pay_one, view_panel)) = true then PANELS_ARRAY(semi_mo_pay_one, view_panel) = PANELS_ARRAY(semi_mo_pay_one, view_panel) * 1		'need to ensure the script is reading and comparing these as numbers
			If IsNumeric(PANELS_ARRAY(semi_mo_pay_two, view_panel)) = true then PANELS_ARRAY(semi_mo_pay_two, view_panel) = PANELS_ARRAY(semi_mo_pay_two, view_panel) * 1
	    End If

	    If PANELS_ARRAY(panel_freq_const, view_panel) = "4 - Weekly" Then
	        If IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = FALSE Then
	            date_to_review = MAXIS_footer_month & "/1/" & MAXIS_footer_year
	            Do
	                If WeekDayName(WeekDay(date_to_review)) = PANELS_ARRAY(panel_weekday_const, view_panel) Then PANELS_ARRAY(panel_known_paydate, view_panel) = date_to_review
	                date_to_review = DateAdd("d", 1, date_to_review)
	            Loop until IsDate(PANELS_ARRAY(panel_known_paydate, view_panel)) = TRUE
	        End If
	    End If
	End If
Next

initial_footer_month = MAXIS_footer_month
initial_footer_year = MAXIS_footer_year
Call navigate_to_MAXIS_screen("STAT", "SUMM")

Do
    this_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
    the_month = DatePart("m", this_month)
    the_year = DatePart("yyyy", this_month)
    next_month = DateAdd("m", 1, this_month)
    next_month_mo = DatePart("m", next_month)
    next_month_yr = DatePart("yyyy", next_month)
    ' Call convert_date_into_MAXIS_footer_month(next_month, next_month_mo, next_month_yr)

    For view_panel = 0 to UBound(PANELS_ARRAY, 2)
        If PANELS_ARRAY(panel_update_checkbox, view_panel) = checked Then
            EMWriteScreen PANELS_ARRAY(panel_type_const, view_panel), 20, 71
            EMWriteScreen PANELS_ARRAY(panel_member_const, view_panel), 20, 76
            EMWriteScreen PANELS_ARRAY(panel_instance_const, view_panel), 20, 79
            transmit

            If PANELS_ARRAY(panel_type_const, view_panel) = "JOBS" Then EMReadScreen end_date, 8, 9, 49
            If PANELS_ARRAY(panel_type_const, view_panel) = "UNEA" Then EMReadScreen end_date, 8, 7, 68
            If end_date = "__ __ __" Then
                PF9
                ' MsgBox "In Edit" & vbNewLine & PANELS_ARRAY(panel_type_const, view_panel) & "-" & PANELS_ARRAY(panel_member_const, view_panel) & "-" & PANELS_ARRAY(panel_instance_const, view_panel) & vbNewLine & PANELS_ARRAY(panel_name_const, view_panel)
            Else
                MsgBox "This panel has an income end date and cannot be updated by the script."
            End If

            'Now that we have read the information we need from the panel as it already exists, we will blank out all the dates and check amounts.
            If PANELS_ARRAY(panel_type_const, view_panel) = "JOBS" Then
                start_row = 12
                end_row = 17
                pay_col = 67
                EMWriteScreen "   ", 18, 43                      'blanking out hours
                EMWriteScreen "   ", 18, 72
            End If
            If PANELS_ARRAY(panel_type_const, view_panel) = "UNEA" Then
                start_row = 13
                end_row = 18
                pay_col = 68
            End If
            panel_row = start_row
            Do
                EMWriteScreen "  ", panel_row, 25            'retro side
                EMWriteScreen "  ", panel_row, 28
                EMWriteScreen "  ", panel_row, 31
                EMWriteScreen "         ", panel_row, pay_col - 29

                EMWriteScreen "  ", panel_row, 54            'prospective side
                EMWriteScreen "  ", panel_row, 57
                EMWriteScreen "  ", panel_row, 60
                EMWriteScreen "         ", panel_row, pay_col

                panel_row = panel_row + 1
            Loop until panel_row = end_row

            check_date = PANELS_ARRAY(panel_known_paydate, view_panel)			'this variable is set from the the known paydate and is changed through the update process here to find each pay date in sequence
            frequency = left(PANELS_ARRAY(panel_freq_const, view_panel), 1)		'identifying the frequency in a variable
            panel_row = start_row												'resetting the panel row for the start of the Do Loop
            total_hours = 0														'resetting the hours count for this partifular panel (source and month specific) to 0 so we can count up as we find paydates in the month

			'comments in this section kept in place for any future bug reports - these are longer message boxes and are ehlpful to have in place.
            Do
                ' MsgBox "Check Date: " & check_date & vbCr & "Month Date Part: " & DatePart("m", check_date) & vbNewLine & "The Month: " & the_month & vbNewLine & vbNewLine & "Year Date Part: " &  DatePart("yyyy", check_date) & vbNewLine & "The Year: " & the_year
                If DatePart("m", check_date) = the_month AND DatePart("yyyy", check_date) = the_year Then		'if the check date is in the current month - we write the information to the panel here.
                    ' MsgBox "MATCH FOUND"
                    call create_mainframe_friendly_date(check_date, panel_row, 54, "YY")						'writing in the information
                    EMWriteScreen PANELS_ARRAY(panel_pay_amt_const, view_panel), panel_row, 68
                    total_hours = total_hours + PANELS_ARRAY(check_hours_const, view_panel)
                    panel_row = panel_row + 1																	'going to the next row
                End If
				'Here we need to go to the next check and that calculation is different depending on the frequency of the pay
                Select case frequency
                    Case "1"													'monthly - just add one month
                        check_date = DateAdd("m", 1, check_date)
                    Case "2"													'semi monthly  - this is the most complicated
                        now_month = DatePart("m", check_date)					'this is the current month
                        now_year = DatePart("yyyy", check_date)
                        later_month = DatePart("m", DateAdd("m", 1, check_date))'this is next month - we need to know this specifically because we need these months when calculating the next paydate
                        later_year = DatePart("yyyy", DateAdd("m", 1, check_date))
						first_of_later_month = later_month & "/1/" & later_year
						' MsgBox "check day - " & DatePart("d", check_date) & vbCr & "semi pay one - " & PANELS_ARRAY(semi_mo_pay_one, view_panel) & vbCr & "semi pay two - " & PANELS_ARRAY(semi_mo_pay_two, view_panel)
						If DatePart("d", check_date) = PANELS_ARRAY(semi_mo_pay_one, view_panel) Then		'if we just added the first paycheck for the month
							If PANELS_ARRAY(semi_mo_pay_two, view_panel) = "Last" Then						'if the second pay always comes on the last day of the month
								check_date = DateAdd("d", -1, first_of_later_month)							'we go one day back from the first day of the next month
                            Else
                                check_date = now_month & "/" & PANELS_ARRAY(semi_mo_pay_two, view_panel) & "/" & now_year		'if the second pay date is a specific date of the month we create the date using the current month and that date
                            End If
                        ElseIf PANELS_ARRAY(semi_mo_pay_two, view_panel) = "Last" Then						'if the date we just added is NOT the first pay date and the second one is always on the last day of the month
							check_date = later_month & "/" & PANELS_ARRAY(semi_mo_pay_one, view_panel) & "/" & later_year		'make the next check the first pay of next month
                        ElseIf DatePart("d", check_date) = PANELS_ARRAY(semi_mo_pay_two, view_panel) Then	'if we just added the second paychek of the month and it ISN'T the last month
							check_date = later_month & "/" & PANELS_ARRAY(semi_mo_pay_one, view_panel) & "/" & later_year		'Make the next check the fist pay of the next month
                        End If
                    Case "3"													'biweekly - just add 14 days
                        check_date = DateAdd("d", 14, check_date)
                    Case "4"													'weekly - just add 7 days
                        check_date = DateAdd("d", 7, check_date)
                End Select
                ' MsgBox check_date
            Loop until DatePart("m", check_date) = next_month_mo AND DatePart("yyyy", check_date) = next_month_yr	'if the next pay check is the next month - we leave the loop because we have all the dates for the current month
            If PANELS_ARRAY(panel_type_const, view_panel) = "JOBS" Then
                total_hours = FormatNumber(total_hours, 0)
                EMWriteScreen "   ", 18, 72
                EMWriteScreen total_hours, 18, 72
            End If
			' MsgBox "Review the panel update because the script thinks it's done."
            Do
                transmit            'save the panel'
                EMReadScreen look_for_warning, 7, 24, 2
            Loop until look_for_warning <> "WARNING"

            ' MsgBox "Look at the updated panel"
        End If
    Next
    MAXIS_footer_month = right("00"&next_month_mo, 2)
    MAXIS_footer_year = right(next_month_yr, 2)
    navigate_there = FALSE

    If MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr Then
        Call back_to_SELF
    Else
        EMWriteScreen "BGTX", 20, 71
        transmit
        EMReadScreen are_we_at_wrap, 4, 2, 46
        If are_we_at_wrap = "WRAP" Then
            EMWriteScreen "Y", 16, 54
            transmit

            EMReadScreen stat_month, 2, 20, 55
            EMReadScreen stat_year, 2, 20, 58

            If stat_month <> MAXIS_footer_month OR stat_year <> MAXIS_footer_year Then navigate_there = TRUE
        Else
            navigate_there = TRUE
        End If

        If navigate_there = TRUE Then
            Call back_to_SELF
            call MAXIS_background_check

            Call navigate_to_MAXIS_screen("STAT", "SUMM")
        End If
    End If

Loop until MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr


script_end_procedure_with_error_report("Complete")
