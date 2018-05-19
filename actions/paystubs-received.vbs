'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PAYSTUBS RECEIVED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 473                	'manual run time in seconds
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("04/23/2018", "Fixed bug in which the lines of the PIC were dupicated in the case note.", "Casey Love, Hennepin County")
CALL changelog_update("12/07/2017", "Removed condition to allow paystubs dated with the current date to be accepted. Updated code to write JOBS verification code in.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/11/2017", "The script has been updated to write to the GRH PIC and to case note that the GRH PIC has been updated.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS
Function prospective_averager(pay_date, gross_amt, hours, paystubs_received, total_prospective_pay, total_prospective_hours) 'Creates variables for total_prospective_pay and total_prospective_hours
  If isdate(pay_date) = True then
    total_prospective_pay = total_prospective_pay + abs(gross_amt)
    total_prospective_hours = total_prospective_hours + abs(hours)
    paystubs_received = paystubs_received + 1
  Else
    pay_date = "01/01/2000"
  End if
End function

Function prospective_pay_analyzer(pay_date, gross_amt)
  If datediff("m", pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0 then
    If len(datepart("m", pay_date)) = 2 then
      EMWriteScreen datepart("m", pay_date), MAXIS_row, 54
    Else
      EMWriteScreen "0" & datepart("m", pay_date), MAXIS_row, 54
    End if
    If len(datepart("d", pay_date)) = 2 then
      EMWriteScreen datepart("d", pay_date), MAXIS_row, 57
    Else
      EMWriteScreen "0" & datepart("d", pay_date), MAXIS_row, 57
    End if
    EMWriteScreen right(datepart("yyyy", pay_date), 2), MAXIS_row, 60
    EMWriteScreen gross_amt, MAXIS_row, 67
    MAXIS_row = MAXIS_row + 1
  End if
End function

Function retro_paystubs_info_adder(pay_date, gross_amt, hours, retro_hours)
  If isdate(pay_date) = True then
    If datediff("m", pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 2 then
      If len(datepart("m", pay_date)) = 2 then
        EMWriteScreen datepart("m", pay_date), MAXIS_row, 25
      Else
        EMWriteScreen "0" & datepart("m", pay_date), MAXIS_row, 25
      End if
      If len(datepart("d", pay_date)) = 2 then
        EMWriteScreen datepart("d", pay_date), MAXIS_row, 28
      Else
        EMWriteScreen "0" & datepart("d", pay_date), MAXIS_row, 28
      End if
      EMWriteScreen right(datepart("yyyy", pay_date), 2), MAXIS_row, 31
      EMWriteScreen gross_amt, MAXIS_row, 38
      retro_hours = abs(retro_hours + abs(hours))
      MAXIS_row = MAXIS_row + 1
    End if
  End if
End function

'DIALOGS----------------------------------------------------------------------------------------------------
'>>>>> This function creates the dialog that the user inputs all the pay stub information.
FUNCTION create_paystubs_received_dialog(worker_signature, number_of_paystubs, paystubs_array, explanation_of_income, employer_name, document_datestamp, pay_frequency, JOBS_verif_code)
	'Declaring the multi-dimensional array for handling pay information
	ReDim paystubs_array(number_of_paystubs - 1, 2)

	BeginDialog paystubs_received_dialog, 0, 0, 256, (160 + (20 * number_of_paystubs - 1)), "Paystubs Received Dialog"
	  DropListBox 100, 5, 100, 15, "(select one)"+chr(9)+"One Time Per Month"+chr(9)+"Two Times Per Month"+chr(9)+"Every Other Week"+chr(9)+"Every Week", pay_frequency
	  FOR i = 0 TO (number_of_paystubs - 1)
		EditBox 15, (45 + (i * 20)), 65, 15, paystubs_array(i, 0)
		EditBox 95, (45 + (i * 20)), 65, 15, paystubs_array(i, 1)
		EditBox 175, (45 + (i * 20)), 65, 15, paystubs_array(i, 2)
	  NEXT
	  EditBox 55, (80 + (20 * (number_of_paystubs - 1))), 190, 15, explanation_of_income
	  EditBox 95, (100 + (20 * (number_of_paystubs - 1))), 80, 15, document_datestamp
	  DropListBox 75, (120 + (20 * (number_of_paystubs - 1))), 120, 15, "(select one)"+chr(9)+"1 Pay Stubs/Tip Report"+chr(9)+"2 Empl Statement"+chr(9)+"3 Coltrl Stmt"+chr(9)+"4 Other Document"+chr(9)+"5 Pend Out State Verification"+chr(9)+"N No Ver Prvd", JOBS_verif_code
	  EditBox 75, (140 + (20 * (number_of_paystubs - 1))), 115, 15, worker_signature
	  ButtonGroup buttonpressed
	  	OkButton 200, (120 + (20 * (number_of_paystubs - 1))), 50, 15
	  	CancelButton 200, (140 + (20 * (number_of_paystubs - 1))), 50, 15
	  Text 40, 10, 55, 10, "Pay frequency:"
	  Text 10, 30, 80, 10, "Pay date (MM/DD/YY):"
	  Text 105, 30, 50, 10, "Gross amount:"
	  Text 195, 30, 30, 10, "Hours:"
	  GroupBox 5, (70 + (20 * (number_of_paystubs - 1))), 245, 30, "Explain how income was calculated:"
	  Text 10, (85 + (20 * (number_of_paystubs - 1))), 45, 10, "Explanation:"
	  Text 10, (105 + (20 * (number_of_paystubs - 1))), 80, 10, "Date paystubs received:"
	  Text 10, (125 + (20 * (number_of_paystubs - 1))), 60, 10, "JOBS verif code:"
	  Text 10, (145 + (20 * (number_of_paystubs - 1))), 60, 10, "Worker signature:"
	EndDialog

	DO
		DO
			err_msg = ""
			DIALOG paystubs_received_dialog
				IF ButtonPressed = 0 THEN stopscript
				If pay_frequency = "(select one)" then err_msg = err_msg & vbCr & "* You must select a pay frequency."
				If JOBS_verif_code = "(select one)" then err_msg = err_msg & vbCr & "You must select a JOBS verif code."
				If explanation_of_income = "" then err_msg = err_msg & vbCr & "* You must explain how you calculated this income (ie: ''all paystubs from last 30 days'')"
				FOR i = 0 TO (number_of_paystubs - 1)
					paystubs_array(i, 0) = Trim(paystubs_array(i, 0))
				NEXT
				FOR i = 0 TO (number_of_paystubs - 1)
					If isdate(paystubs_array(i, 0)) = False AND paystubs_array(i, 0) <> "" THEN err_msg = err_msg & vbCr & "* Your pay date must be ''MM/DD/YYYY'' format. Please try again."
				NEXT
				FOR i = 0 TO (number_of_paystubs - 1)
					If isdate(paystubs_array(i, 0)) = True AND datediff("d", date, paystubs_array(i, 0)) > 0 THEN err_msg = err_msg & vbCr & "* You cannot enter a paydate in the future. Please remove and try again."
				NEXT
				FOR i = 0 TO (number_of_paystubs - 1)
					If isdate(paystubs_array(i, 0)) = True and (Isnumeric(paystubs_array(i, 1)) = False or Isnumeric(paystubs_array(i, 2)) = False) then err_msg = err_msg & vbCr & "* You must include a gross pay amount as well as an hours amount."
				NEXT
				FOR i = 0 TO (number_of_paystubs - 1)
					IF paystubs_array(i, 0) = "" THEN err_msg = err_msg & vbCr & "* You cannot leave pay dates blank."
					IF paystubs_array(i, 1) = "" OR paystubs_array(i, 2) = "" THEN err_msg = err_msg & vbCr & "* You cannot leave pay information blank."
				NEXT
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false
END FUNCTION

BeginDialog number_of_paystubs_dlg, 0, 0, 211, 65, "Number of Pay Dates"
  EditBox 165, 10, 40, 15, number_of_paystubs
  ButtonGroup ButtonPressed
    OkButton 105, 45, 50, 15
    CancelButton 155, 45, 50, 15
  Text 10, 15, 145, 10, "Enter the number of pay dates being used..."
EndDialog

BeginDialog paystubs_received_case_number_dialog, 0, 0, 376, 170, "Case number"
  EditBox 100, 5, 60, 15, MAXIS_case_number
  EditBox 70, 25, 25, 15, MAXIS_footer_month
  EditBox 125, 25, 25, 15, MAXIS_footer_year
  EditBox 110, 45, 25, 15, HH_member
  CheckBox 15, 75, 110, 10, "Update and case note the PIC?", update_PIC_check
  CheckBox 15, 90, 75, 10, "Update HC popup?", update_HC_popup_check
  CheckBox 15, 105, 130, 10, "Update and case note the GRH PIC?", update_GRH_PIC_check
  CheckBox 15, 120, 140, 10, "Check here to have the script update all", future_months_check
  CheckBox 15, 145, 135, 10, "Case note info about paystubs?", add_case_note_check
  ButtonGroup ButtonPressed
    OkButton 265, 150, 50, 15
    CancelButton 320, 150, 50, 15
  Text 10, 10, 85, 10, "Enter your case number:"
  GroupBox 175, 5, 195, 140, "INSTRUCTIONS!!! PLEASE READ!!!"
  Text 185, 20, 180, 35, "This script, by default, will update retro/pro in the footer month specified only. It can update multiple months and send through background if you select that to the left. It can also update the PIC or HC pop-ups."
  Text 185, 60, 180, 50, "PLEASE NOTE: you should already have a JOBS panel made for this client. If you haven't made a JOBS panel yet, make it and send the case through background before using this script. The script only does one job at a time, so you may need to run it more than once if you have multiple jobs."
  Text 185, 115, 180, 25, "You should also have all of the paystubs you need to update MAXIS. If you aren't ready to update STAT/JOBS, don't use this script."
  Text 20, 30, 50, 10, "Footer month:"
  Text 100, 30, 20, 10, "Year:"
  Text 35, 50, 75, 10, "HH memb # for JOBS:"
  GroupBox 10, 65, 150, 95, "Options"
  Text 30, 130, 120, 10, "future months and send through BG."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Default member is member 01
HH_member = "01"

DO
	'Shows the case number dialog
	DO
		Dialog paystubs_received_case_number_dialog
			If buttonpressed = 0 then stopscript
		call check_for_password(are_we_passworded_out)
	LOOP UNTIL are_we_passworded_out = false

	CALL check_for_MAXIS(False)							'checkng for an active MAXIS session
	Call MAXIS_footer_month_confirmation				'confirming that user is in the correct footer month
	call navigate_to_MAXIS_screen("stat", "jobs")		'navigates to STAT/JOBS

	'Heads into the case/curr screen, checks to make sure the case number is correct before proceeding. If it can't get beyond the SELF menu the script will stop.
	EMReadScreen SELF_check, 4, 2, 50
	If SELF_check = "SELF" then stopscript

	'Navigates to the JOBS panel for the right person
	If HH_member <> "01" then
		EMWriteScreen HH_member, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
	End if

	'Checks to make sure there are JOBS panels for this member. If none exist the script will close
	EMReadScreen total_amt_of_panels, 1, 2, 78
	If total_amt_of_panels = "0" then script_end_procedure("No JOBS panels exist for this client. Please add a JOBS panel and run through background before trying again. The script will now stop.")

	'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
	If total_amt_of_panels <> "0" then
		Do
			EMReadScreen current_panel_number, 1, 2, 73
			EMReadScreen employer_name, 30, 7, 42
			employer_check = MsgBox("Is this your employer? Employer name: " & trim(replace(employer_name, "_", "")), 3)
			If employer_check = 2 then stopscript
			If employer_check = 6 then
				employer_found = True
				exit do
			END IF
			If employer_check = 7 and current_panel_number = total_amt_of_panels then
				employer_found = False
				pick_a_different_household_member = MsgBox("You have run through all the possible employers for this person. If you need to select a different household member, press OK. If you need to stop the script to change the case number, create a new job, etc, press CANCEL.", vbOKCancel)
				IF pick_a_different_household_member = vbCancel THEN stopscript
				IF pick_a_different_household_member = vkOK THEN EXIT DO
			End if
			transmit
		Loop until current_panel_number = total_amt_of_panels
	End if

	'Reads employer name for case note
	EMReadScreen employer_name, 30, 7, 42			'Read the name
	employer_name = replace(employer_name, "_", "")		'Clean up the name with replacing underscores
	call fix_case(employer_name, 3)				'and using custom fix_case function to set case

LOOP UNTIL employer_found = True

DO
	DIALOG number_of_paystubs_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF IsNumeric(number_of_paystubs) = False THEN MsgBox "Please enter the number of pay dates as a number."
LOOP UNTIL ButtonPressed = -1 AND number_of_paystubs <> "" AND IsNumeric(number_of_paystubs) = True

'Shows the paystub dialog. Includes logic to prevent paydates from being entered incorrectly.

DO
	CALL create_paystubs_received_dialog(worker_signature, number_of_paystubs, paystubs_array, explanation_of_income, employer_name, document_datestamp, pay_frequency, JOBS_verif_code)

		'From Robert Fewins-Kalb on 03/31/2016 (because this script is massive and I want to document want what is added, when and why)
		'Making sure that the script returns to the JOBS panel it is supposed to.
		'Addition of check_for_password appears to force through a transmit which causes the script to jump ahead 1 JOBS panel
	CALL write_value_and_transmit("0" & current_panel_number, 20, 79)
	'Do
	'    EMReadScreen selected_employer, 30, 7, 42
	'    If trim(selected_employer) <> trim(employer_name) then
	'		correct_job = False
	'		confirm_job = msgbox("The employer you selected to update was " & employer_name. " This JOBS panel does not match. Either add a panel now, or navigate to the correct job. Press OK when ready for script to continue.", vbOkCancel + vbExclamation, "JOBS panel names do not match.")
	'    	If confirm_job = vbCancel then script_end_procedure("You have chosen to stop the script. Please review case for incomplete changes or inforamtion.")
	'		IF confirm_job = vbOK then correct_job = true
	'	End if
	'Loop until correct_job = true

	err_msg = ""
	'Checking dates to make sure all are on the same day of the week, in instances of weekly or biweekly income. This avoids a possible issue
	'resulting from a paydate being "moved" due to a holiday, and affecting the rest of the calculation for income. If the dates are not all on the
	'same day, the script will stop.
	If pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then
		weekday_baseline = weekday(cdate(paystubs_array(0, 0)))
		list_of_weekdays = "Select one..."
		list_of_weekdays = list_of_weekdays+chr(9)+WeekDayName(weekday(cdate(paystubs_array(0, 0))))

		FOR i = 1 TO (number_of_paystubs - 1)
			IF paystubs_array(i, 0) <> "" THEN
				IF WeekDay(CDate(paystubs_array(i, 0))) <> weekday_baseline THEN
					err_msg = err_msg & vbCr & (paystubs_array(i, 0) & " is on a different pay date than the first pay date on the script.")
					IF InStr(list_of_weekdays, WeekDayName(weekday(cdate(paystubs_array(i, 0))))) = 0 THEN list_of_weekdays = list_of_weekdays+chr(9)+WeekDayName(weekday(cdate(paystubs_array(i, 0))))
				END IF
			END IF
		NEXT
	END IF

	IF err_msg <> "" THEN
		dates_not_aligned = MsgBox("*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "The script is going to ask you to pick a specific day of the week to use for prospecting income. Press OK to continue. If you do not want to do this, press CANCEL (the script will end).", vbOKCancel)
		IF dates_not_aligned = vbCancel THEN script_end_procedure("")

        BeginDialog weekday_dlg, 0, 0, 191, 150, "Pick a Weekday"
          Text 15, 15, 160, 35, "Pick a specific day of the week for the script to use for prospecting income. The script is using the days of the week from the pay stubs you entered in the previous dialog."
          Text 15, 70, 40, 10, "Weekday:"
          DropListBox 65, 70, 70, 15, list_of_weekdays, weekday_to_use
          ButtonGroup ButtonPressed
            OkButton 85, 130, 50, 15
            CancelButton 135, 130, 50, 15
        EndDialog

		DO
			err_msg = ""
			DIALOG weekday_dlg
				IF ButtonPressed = 0 THEN stopscript
				IF weekday_to_use = "Select one..." THEN MsgBox "Select a weekday."
		LOOP UNTIL weekday_to_use <> "Select one..."
	ELSE
		weekday_to_use = WeekDayName(WeekDay(paystubs_array(0, 0)))
	END IF
LOOP UNTIL err_msg = ""

'Turns on edit mode
PF9

'Declares variables it'll need for the next part
dim paystubs_received
dim total_prospective_pay
dim total_prospective_hours

'Totals the prospective amounts, inserts "01/01/2000" for dates that were left blank, using function.
FOR i = 0 TO (number_of_paystubs - 1)
	Call prospective_averager(paystubs_array(i, 0), paystubs_array(i, 1), paystubs_array(i, 2), paystubs_received, total_prospective_pay, total_prospective_hours)
NEXT


'Creates averages
average_pay_per_paystub = formatnumber(total_prospective_pay / paystubs_received, 2, 0, 0, 0)
average_hours_per_paystub = abs(total_prospective_hours / paystubs_received)


Do
	'If SNAP was active the script must update the PIC.
	If update_PIC_check = 1 then
		IF number_of_paystubs > 10 THEN
			MsgBox "You indicated you are using more than 10 pay dates. The PIC cannot handle more than 10 pay dates. The script will not be able to update the PIC. You will need to process the PIC manually."
		ELSE
			EMWriteScreen "x", 19, 38
			transmit

			'Determining if there is a page 2 on the PIC
			PF20
			EMReadScreen complete_the_page, 17, 20, 6
			IF complete_the_page <> "COMPLETE THE PAGE" THEN
				FOR a = 9 to 13
					EMSetCursor a, 13
					EMSendKey "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"
				NEXT
				PF19
				PF19
			END IF
			'Clears existing info off PIC
			EMSendKey "<home>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>" + "<tab>" + "<eraseeof>"

			'The following will generate a MAXIS formatted date for today.
			current_day = DatePart("D", date)
			If len(current_day) = 1 then current_day = "0" & current_day
			current_month = DatePart("M", date)
			If len(current_month) = 1 then current_month = "0" & current_month
			current_year = right(DatePart("yyyy", date), 2)
			'Puts current date and pay frequency in PIC.
			CALL create_MAXIS_friendly_date(date, 0, 5, 34)
			If pay_frequency = "One Time Per Month" then EMWriteScreen "1", 5, 64
			If pay_frequency = "Two Times Per Month" then EMWriteScreen "2", 5, 64
			If pay_frequency = "Every Other Week" then EMWriteScreen "3", 5, 64
			If pay_frequency = "Every Week" then EMWriteScreen "4", 5, 64
			'Sets PIC row for the next functions
			DIM PIC_row
			PIC_row = 9
			'Uses function to add each PIC pay date, income, and hours. Doesn't add any if they show "01/01/2000" as those are dummy numbers
			FOR i = 0 to (number_of_paystubs - 1)
				IF paystubs_array(i, 0) <> "01/01/2000" THEN
					If isdate(paystubs_array(i, 0)) = True then
						CALL create_MAXIS_friendly_date(paystubs_array(i, 0), 0, PIC_row, 13)
						EMWriteScreen paystubs_array(i, 1), PIC_row, 25
						EMWriteScreen paystubs_array(i, 2), PIC_row, 35
						PIC_row = PIC_row + 1
						IF PIC_row = 14 THEN
							PF20			'navigates to page 2 of the PIC'
							PF20
							PIC_row = 9
						ELSE
							Transmit		'Transmits in order to format the PIC, but cannot do this if page 2 of the PIC is going to be populated
							Transmit		'It would just bring you back to the JOBS panel'
						END IF
					End If
				END IF
			NEXT

			'Reading the PIC if update_PIC_check was checked
			PF19 									'navigates to the 1st page of the PIC (even if there's only one PIC)'
			'Reads the contents of the PIC for case noting.
			EMReadScreen PIC_line_01, 26, 5, 49
			EMReadScreen PIC_line_02, 28, 8, 13
			EMReadScreen PIC_line_03, 28, 9, 13
			EMReadScreen PIC_line_04, 28, 10, 13
			EMReadScreen PIC_line_05, 28, 11, 13
			EMReadScreen PIC_line_06, 28, 12, 13
			EMReadScreen PIC_line_07, 28, 13, 13
			EMReadScreen PIC_line_08, 28, 14, 13
			EMReadScreen PIC_line_09, 50, 16, 22
			EMReadScreen PIC_line_10, 50, 17, 22
			EMReadScreen PIC_line_11, 50, 18, 22
            If PIC_line_07 <> "__ __ __    ________  ______" Then
    			PF20										'shift PF8 to the next PIC screen'
    			EMReadScreen PIC_page_2_check, 8, 20, 6
                MagBox PIC_page_2_check
    			IF PIC_page_2_check <> "COMPLETE" THEN
    				EMReadScreen PIC2_line_01, 28, 9, 13		'reading the 2nd page of the PIC'
                    If PIC2_line_01 <> PIC_line_03 Then
        				EMReadScreen PIC2_line_02, 28, 10, 13
        				EMReadScreen PIC2_line_03, 28, 11, 13
        				EMReadScreen PIC2_line_04, 28, 12, 13
        				EMReadScreen PIC2_line_05, 28, 13, 13
        				PIC2_line_01 = Replace(PIC2_line_01, "$", " ")
        				PIC2_line_01 = Replace(PIC2_line_01, "__ __ __    ________  ______", "")
        				PIC2_line_02 = Replace(PIC2_line_02, "__ __ __    ________  ______", "")
        				PIC2_line_03 = Replace(PIC2_line_03, "__ __ __    ________  ______", "")
        				PIC2_line_04 = Replace(PIC2_line_04, "__ __ __    ________  ______", "")
        				PIC2_line_05 = Replace(PIC2_line_05, "__ __ __    ________  ______", "")
                    Else
                        PIC2_line_01 = ""
                    End If
    			END IF
            End If
			transmit
		END IF
    End if

	'going into the GRH PIC to update...
	IF update_GRH_PIC_check = 1 THEN
		'checking to make sure that the user has the case in a benefit month that includes the GRH PIC... 07/16 is the first month...
		EMReadScreen grh_pic, 7, 19, 73
		IF grh_pic <> "GRH PIC" THEN
			MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "You are attempting to update the GRH PIC in a budget month prior to the implementation of the GRH PIC on STAT/JOBS. The script will skip attempting to update the GRH PIC for this month.", vbExclamation
		ELSE
			'else, going in to the GRH PIC
			CALL write_value_and_transmit("X", 19, 71)

			'erasing the information currently in the GRH PIC
			EMWriteScreen "_", 3, 63		'pay frequency
			EMWriteScreen "______", 6, 63		'hrs/wk
			EMWriteScreen "________", 7, 65		'rate/hr
			EMWriteScreen "________", 11, 65	'non-recurring
			FOR row = 7 to 16
				EMWriteScreen "__", row, 9
				EMWriteScreen "__", row, 12
				EMWriteScreen "__", row, 15
				EMWriteScreen "________", row, 21
			NEXT

			'writing today's date in the Date of Calculation field
			CALL create_mainframe_friendly_date(date, 3, 30, "YY")

			'writing the pay frequency
			If pay_frequency = "One Time Per Month" then 	EMWriteScreen "1", 3, 63
			If pay_frequency = "Two Times Per Month" then 	EMWriteScreen "2", 3, 63
			If pay_frequency = "Every Other Week" then 		EMWriteScreen "3", 3, 63
			If pay_frequency = "Every Week" then 			EMWriteScreen "4", 3, 63

			'updating income lines
			GRH_PIC_row = 7
			'Uses function to add each PIC pay date, income, and hours. Doesn't add any if they show "01/01/2000" as those are dummy numbers
			FOR i = 0 to (number_of_paystubs - 1)
				IF paystubs_array(i, 0) <> "01/01/2000" THEN
					If isdate(paystubs_array(i, 0)) = True then
						CALL create_mainframe_friendly_date(paystubs_array(i, 0), GRH_PIC_row + i, 9, "YY")
						EMWriteScreen paystubs_array(i, 1), GRH_PIC_row + i, 21
					End If
				END IF
			NEXT
			transmit
			EMReadScreen avg_grh_income, 39, 16, 38
			EMReadScreen grh_prosp_monthly, 42, 17, 35
			PF3
		END IF
	END IF

	'Clears JOBS data before updating the JOBS panel
	EMSetCursor 12, 25
	EMSendKey "___________________________________________________________________________________________________________________________________________________"

	'Updates for retrospective income by checking each pay date's month against the footer month using a function. If the footer month is two months ahead of the pay month it will add to JOBS and keep a tally of hours.
	MAXIS_row = 12 'Needs this for the following functions
	Dim retro_hours
	FOR i = 0 TO (number_of_paystubs - 1)
		CALL retro_paystubs_info_adder(paystubs_array(i, 0), paystubs_array(i, 1), paystubs_array(i, 2), retro_hours)
	NEXT

	'Must convert retro hours into an integer for MAXIS
	retro_hours = retro_hours + .00000000000001 'This will force rounding to go half-up, as the CINT function rounds half down, which goes against procedure.
	retro_hours = cint(retro_hours)

	'Puts hours worked in the retro months in. This was determined using the previous functions.
	If retro_hours > 999 then retro_hours = 999 'In case there are over 999 hours, this is the procedure
	If retro_hours <> "" and retro_hours <> 0 then EMWriteScreen retro_hours, 18, 43
	retro_hours = 0 'Clears variable so it can be used in multiple months if needed

	'Determines the paydate to put in the prospective side. It moves forward for instances where the footer month is ahead of the first paydate, otherwise it moves backward until it lands on the right date.
	first_prospective_pay_date = ""
	FOR i = 0 TO (number_of_paystubs - 1)
		IF WeekDayName(WeekDay(paystubs_array(i, 0))) = weekday_to_use THEN
			IF first_prospective_pay_date = "" THEN
				first_prospective_pay_date = paystubs_array(i, 0)
			ELSE
				'If the paystubs_array(i, 0) is earlier than the existing first_prospective_pay_date THEN the script resets first_prospective_pay_date with the value of paystubs_array(i, 0)
				IF DateDiff("D", first_prospective_pay_date, paystubs_array(i, 0)) < 0 THEN first_prospective_pay_date = paystubs_array(i, 0)
			END IF
		END IF
	NEXT

	If datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) > 0 then 'For instances where the footer month is ahead of the first paydate.
		Do
			If datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0 then exit do
			If pay_frequency = "One Time Per Month" then first_prospective_pay_date = dateadd("m", 1, first_prospective_pay_date)
			If pay_frequency = "Two Times Per Month" then first_prospective_pay_date = dateadd("m", 1, first_prospective_pay_date)
			If pay_frequency = "Every Other Week" then first_prospective_pay_date = dateadd("d", 14, first_prospective_pay_date)
			If pay_frequency = "Every Week" then first_prospective_pay_date = dateadd("d", 7, first_prospective_pay_date)
		Loop until datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0
	Elseif datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) < 0 then 'For instances where the footer month is behind the first paydate (ex: paydate is 06/26/2013 but footer month is 05/13).
		Do
			If datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0 then exit do
			If pay_frequency = "One Time Per Month" then first_prospective_pay_date = dateadd("m", -1, first_prospective_pay_date)
			If pay_frequency = "Two Times Per Month" then first_prospective_pay_date = dateadd("m", -1, first_prospective_pay_date)
			If pay_frequency = "Every Other Week" then first_prospective_pay_date = dateadd("d", -14, first_prospective_pay_date)
			If pay_frequency = "Every Week" then first_prospective_pay_date = dateadd("d", -7, first_prospective_pay_date)
		Loop until datediff("m", first_prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0
	End if
	'This checks to make sure the earliest possible paydate is selected in each prospective month.
	If pay_frequency = "Two Times Per Month" or pay_frequency = "Every Other Week" or pay_frequency = "Every Week" then
		Do
			If pay_frequency = "Two Times Per Month" and datepart("d", first_prospective_pay_date) > 15 then first_prospective_pay_date = dateadd("d", -15, first_prospective_pay_date)
			If pay_frequency = "Every Other Week" and datepart("d", first_prospective_pay_date) > 14 then first_prospective_pay_date = dateadd("d", -14, first_prospective_pay_date)
			If pay_frequency = "Every Week" and datepart("d", first_prospective_pay_date) > 7 then first_prospective_pay_date = dateadd("d", -7, first_prospective_pay_date)
		Loop until (pay_frequency = "Two Times Per Month" and datepart("d", first_prospective_pay_date) <= 15) or (pay_frequency = "Every Other Week" and datepart("d", first_prospective_pay_date) <= 14) or (pay_frequency = "Every Week" and datepart("d", first_prospective_pay_date) <= 7)
	End if


	'Analyzes the paystubs received using a function, puts any actual paystubs received in the footer month into the JOBS panel on the prospective side.
	MAXIS_row = 12 'This variable is needed for the script to know which line to put the prospective info on
	FOR i = 0 TO (number_of_paystubs - 1)
		CALL prospective_pay_analyzer(paystubs_array(i, 0), paystubs_array(i, 1))
	NEXT
	total_prospective_dates = MAXIS_row - 12

	'Adds the remaining weeks in using a do...loop to determine all of the anticipated pay dates for the client.
	If pay_frequency = "One Time Per Month" then pay_multiplier = 31
	If pay_frequency = "Two Times Per Month" then pay_multiplier = 15
	If pay_frequency = "Every Other Week" then pay_multiplier = 14
	If pay_frequency = "Every Week" then pay_multiplier = 7

	Do
		If pay_frequency = "One Time Per Month" and total_prospective_dates >= 1 then exit do 'Shouldn't be more than one entry if pay is once per month.
		If pay_frequency = "Two Times Per Month" and total_prospective_dates >= 2 then exit do 'Shouldn't be more than two entries if pay is twice per month.
		prospective_pay_date = dateadd("d", total_prospective_dates * pay_multiplier, first_prospective_pay_date)
		If datediff("m", prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 0 then
			If len(datepart("m", prospective_pay_date)) = 2 then
				EMWriteScreen datepart("m", prospective_pay_date), MAXIS_row, 54
			Else
				EMWriteScreen "0" & datepart("m", prospective_pay_date), MAXIS_row, 54
			End if
			If len(datepart("d", prospective_pay_date)) = 2 then
				EMWriteScreen datepart("d", prospective_pay_date), MAXIS_row, 57
			Else
				EMWriteScreen "0" & datepart("d", prospective_pay_date), MAXIS_row, 57
			End if
			EMWriteScreen right(datepart("yyyy", prospective_pay_date), 2), MAXIS_row, 60
			EMWriteScreen average_pay_per_paystub, MAXIS_row, 67
			MAXIS_row = MAXIS_row + 1
			total_prospective_dates = total_prospective_dates + 1
		End if
	Loop until datediff("m", prospective_pay_date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) <> 0
	'Updates pay frequency
	If pay_frequency = "One Time Per Month" then EMWriteScreen "1", 18, 35
	If pay_frequency = "Two Times Per Month" then EMWriteScreen "2", 18, 35
	If pay_frequency = "Every Other Week" then EMWriteScreen "3", 18, 35
	If pay_frequency = "Every Week" then EMWriteScreen "4", 18, 35

	'Puts average hours in. Added a small imperfection ".0000000000001" so that if any hourly amounts land on exactly ".5", they will round half-up instead of half down.
	If pay_frequency = "One Time Per Month" then EMWriteScreen cint(average_hours_per_paystub + .0000000000001), 18, 72
	If pay_frequency = "Two Times Per Month" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72
	If pay_frequency = "Every Other Week" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72
	If pay_frequency = "Every Week" then EMWriteScreen cint((average_hours_per_paystub + .0000000000001) * total_prospective_dates), 18, 72

	'Puts pay verification type in. JOBS panel verification codes were updated in 10-16 to coordinates 6, 34 from coordinates 6, 38
	EMWriteScreen left(JOBS_verif_code, 1), 6, 34

	'If the footer month is the current month + 1, the script needs to update the HC popup for HC cases.
	If update_HC_popup_check = 1 and datediff("m", date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 1 then
		EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
		IF HC_income_est_check = "Est" Then 'this is the old position
			EMWriteScreen "x", 19, 54
		ELSE								'this is the new position
			EMWriteScreen "x", 19, 48
		END IF
		transmit
		EMWriteScreen "________", 11, 63
		EMWriteScreen average_pay_per_paystub, 11, 63
		Do 'Doing this as a pop-up since there are times when a warning message changes the amount of times this plays.
			transmit
			EMReadScreen HC_popup_check, 18, 9, 43
			If HC_popup_check <> "HC Income Estimate" then updated_HC_popup = True
		Loop until HC_popup_check <> "HC Income Estimate"
	End if

	'Transmits after ending the JOBS panel updating
	Do
		transmit
		EMReadScreen display_mode_check, 1, 20, 8
	Loop until display_mode_check = "D"

	If datediff("m", date, MAXIS_footer_month & "/01/" & MAXIS_footer_year) = 1 then in_future_month = True

	'If just on SNAP, the case does not have to update future months, so the script can now case note.
	If future_months_check = 0 or in_future_month = True then exit do

	'Navigates to the current month + 1 footer month, then back into the JOBS panel
	CALL write_value_and_transmit("BGTX", 20, 71)
	CALL write_value_and_transmit("y", 16, 54)
	EMReadScreen MAXIS_footer_month, 2, 20, 55
	EMReadScreen MAXIS_footer_year, 2, 20, 58
	EMWriteScreen "jobs", 20, 71
	EMWriteScreen HH_member, 20, 76
	If len(current_panel_number) = 1 then current_panel_number = "0" & current_panel_number
	EMWriteScreen current_panel_number, 20, 79
	transmit
	PF9
Loop until in_future_month = True

'Case noting section
If update_PIC_check = 1 then
	start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("~~~SNAP PIC for MEMB " & HH_member & ": " & date & "~~~")
	EMSendKey PIC_line_02 & "<newline>"
	EMSendKey PIC_line_03 & "                 " & "<newline>"
	EMSendKey PIC_line_04 & "                 " & "<newline>"
	EMSendKey PIC_line_05 & "                 " & "<newline>"
	EMSendKey PIC_line_06 & "                 " & "<newline>"
	EMSendKey PIC_line_07 & "                 " & "<newline>"
	IF PIC2_line_01 <> "" then EMSendKey PIC2_line_01 & "                 " & "<newline>"
	IF PIC2_line_02 <> "" then EMSendKey PIC2_line_02 & "                 " & "<newline>"
	IF PIC2_line_03 <> "" then EMSendKey PIC2_line_03 & "                 " & "<newline>"
	IF PIC2_line_04 <> "" then EMSendKey PIC2_line_04 & "                 " & "<newline>"
	IF PIC2_line_05 <> "" then EMSendKey PIC2_line_05 & "                 " & "<newline>"
	EMSendKey PIC_line_08 & "<newline>"
	EMWriteScreen PIC_line_01, 6, 48
	EMWriteScreen PIC_line_09, 7, 35
	EMWriteScreen PIC_line_10, 8, 35
	EMWriteScreen PIC_line_11, 9, 35
	If explanation_of_income <> "" then
		EMSendKey "---" & "<newline>"
		call write_bullet_and_variable_in_CASE_NOTE("How income was calculated", explanation_of_income)
	End if
	call write_bullet_and_variable_in_CASE_NOTE("Employer name", employer_name)
	If document_datestamp <> "" then call write_bullet_and_variable_in_CASE_NOTE("Paystubs received date", document_datestamp)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
End if

IF update_GRH_PIC_check = 1 THEN
	start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("~~~GRH PIC for MEMB " & HH_member & ": " & date & "~~~")
	CALL write_variable_in_CASE_NOTE("Pay Date    Gross Amt")
	FOR i = 0 TO (number_of_paystubs - 1)
		CALL write_variable_in_CASE_NOTE(paystubs_array(i, 0) & "    " & FormatCurrency(paystubs_array(i, 1)))
	NEXT
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(avg_grh_income)
	CALL write_variable_in_CASE_NOTE(grh_prosp_monthly)
	CALL write_variable_in_CASE_NOTE("---")
	call write_bullet_and_variable_in_CASE_NOTE("How income was calculated", explanation_of_income)
	call write_bullet_and_variable_in_CASE_NOTE("Employer name", employer_name)
	If document_datestamp <> "" then call write_bullet_and_variable_in_CASE_NOTE("Paystubs received date", document_datestamp)
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

If add_case_note_check = 1 then
	start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("Paystubs received for MEMB " & HH_member & ": updated JOBS w/script")
	call write_three_columns_in_case_note(14, "DATE", 29, "AMT", 39, "HOURS")
	FOR i = 0 TO (number_of_paystubs - 1)
		IF paystubs_array(i, 0) <> "01/01/2000" THEN CALL write_three_columns_in_case_note(12, paystubs_array(i, 0), 27, "$" & paystubs_array(i, 1), 39, paystubs_array(i, 2))
	NEXT
	If explanation_of_income <> "" then
		EMSendKey "---" & "<newline>"
		call write_bullet_and_variable_in_CASE_NOTE("How income was calculated", explanation_of_income)
	End if
	call write_bullet_and_variable_in_CASE_NOTE("Employer name", employer_name)
	If document_datestamp <> "" then call write_bullet_and_variable_in_CASE_NOTE("Paystubs received date", document_datestamp)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
End if

IF number_of_paystubs > 5 THEN
	MsgBox "Success!! Your JOBS panel has been updated. However, because you have used more than 5 pay dates, the script may not have updated the retro side appropriately, please double check the retro side of your JOBS panel. You may need to manually update it to get the pay information updated correctly."
ELSE
	MsgBox "Success!! Your JOBS panel has been updated with the paystubs you've entered in. Send your case through background, review the results, and take action as appropriate. Don't forget to case note!"
END IF
script_end_procedure("")
