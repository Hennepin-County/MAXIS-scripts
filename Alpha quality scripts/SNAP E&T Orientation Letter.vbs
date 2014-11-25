'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - FSET "
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------



'Main Dialog
BeginDialog FSET_dialog, 0, 0, 176, 145, "FSET appointment letter"
  ButtonGroup ButtonPressed
    OkButton 45, 125, 50, 15
    CancelButton 105, 125, 50, 15
  EditBox 85, 10, 70, 15, case_number
  EditBox 85, 30, 20, 15, member_number
  EditBox 85, 55, 70, 15, appointment_date
    Text 5, 10, 45, 15, "Case Number"
  Text 5, 55, 45, 15, "Appointment Date"
  Text 5, 100, 105, 20, "Sign your case note."
  Text 5, 75, 70, 15, "FSET Provider Location"
  DropListBox 85, 80, 70, 20, "Duluth"+chr(9)+"Hibbing"+chr(9)+"Virginia", FSET_provider
  EditBox 85, 100, 70, 15, worker_signature
  Text 5, 30, 65, 15, "Member Number"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""

'Shows dialog, checks for password prompt
  Do
    Do
      Dialog FSET_dialog
      If ButtonPressed = 0 then stopscript
      If case_number = "" then MsgBox "You did not enter a case number. Please try again."
      If appointment_date = "" then MsgBox "You did not enter an appointment date. Please try again."
      If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
	  If FSET_provider = "" then MsgBox "You did not select an FSET provider Location.  Please try again."
	  If member_number = "" then MsgBox "You did not specify a household member.  Please try again."
	  If isdate(appointment_date) = False then MsgBox "You did not enter a valid appointment date. Please try again."
	
	Loop until case_number <> "" and appointment_date <> "" and worker_signature <> "" and FSET_provider <> "" and member_number <> "" and isdate(appointment_date) = True
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be outside of MAXIS. You may be locked out of MAXIS, check your screen and try again."
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

  'Pulls the member name.
 call navigate_to_screen("STAT", "MEMB")
 EMWriteScreen member_number, 20, 76
 transmit
 EMReadScreen last_name, 24, 6, 30
 EMReadScreen first_name, 11, 6, 63
 last_name = trim(replace(last_name, "_", ""))
 first_name = trim(replace(first_name, "_", ""))
 
 'Navigates into SPEC/LETR
  call navigate_to_screen("SPEC", "LETR")
  
  'Checks to make sure we're past the SELF menu
  EMReadScreen still_self, 27, 2, 28 
  If still_self = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past the SELF screen. Is your case in background?")
  
  'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
  EMWriteScreen "x", 8, 12
  transmit
  EMReadScreen LETR_check, 4, 2, 49
  If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

  'Sets the info variables for the letter.  Each city has 1 provider that uses set appointment times in SLC.
  If FSET_provider = "Duluth" then
     time_hour = "09"
	 time_minute = "30"
	 time_ampm = "AM"
	 loca_line1 = "Minnesota Workforce Center Duluth"
	 loca_line2 = "402 W 1st St Room 119"
	 loca_line3 = "Duluth MN 55802"
	 phone_1 = "302"
	 phone_2 = "8400"
  End If
  If FSET_provider = "Virginia" then
     time_hour = "01"
	 time_minute = "30"
	 time_ampm = "PM"
	 loca_line1 = "Virginia Workforce Center"
	 loca_line2 = "820 9th St"
	 loca_line3 = "Virginia MN 55792"
	 phone_1 = "471"
	 phone_2 = "7515"
  End If
  If FSET_provider = "Hibbing" then
     time_hour = "01"
	 time_minute = "00"
	 time_ampm = "PM"
	 loca_line1 = "Minnesota Workforce Center Hibbing"
	 loca_line2 = "3920 13th Avenue E"
	 loca_line3 = "Hibbing MN 55746"
	 phone_1 = "231"
	 phone_2 = "8590"
  End If
  
 
 'Creates a Maxis friendly date for spacing of date fields on the FSET letter screen (which are different from most maxis screens)
 Function create_FSET_friendly_date(date_variable, variable_length, screen_row, screen_col) 
  var_month = datepart("m", dateadd("d", variable_length, date_variable))
  If len(var_month) = 1 then var_month = "0" & var_month
  EMWriteScreen var_month, screen_row, screen_col
  var_day = datepart("d", dateadd("d", variable_length, date_variable))
  If len(var_day) = 1 then var_day = "0" & var_day
  EMWriteScreen var_day, screen_row, screen_col + 5
  var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
  EMWriteScreen right(var_year, 2), screen_row, screen_col + 10
End function


  'Writes the info into the LETR. 
  EMWriteScreen first_name & " " & last_name, 4, 28
  call create_FSET_friendly_date(appointment_date, 0, 6, 28) 
  EMWriteScreen time_hour, 7, 28
  EMWriteScreen time_minute, 7, 33
  EMWriteScreen time_ampm, 7, 38
  EMWriteScreen loca_line1, 9, 28
  EMWriteScreen loca_line2, 10, 28
  EMWriteScreen loca_line3, 11, 28
  EMWriteScreen "218", 13, 28
  EMWriteScreen phone_1, 13, 34
  EMWriteScreen phone_2, 13, 40
  EMWriteScreen "AEOA", 16, 28
  transmit
  PF4
      
 'Navigates to a blank case note
  call navigate_to_screen("case", "note")
  PF9
    
 'Writes the case note
  EMSendKey "**FSET Appointment letter sent**" & "<newline>"
  EMSendKey "FSET Appointment was scheduled for " & appointment_date & " at " & time_hour & ":" &  time_minute & time_ampm & "<newline>" 
  EMSendKey "At AEOA: " & loca_line1 & "<newline>"
  EMSendKey worker_signature
  MsgBox "Success! The FSET Letter has been sent and a case note has been made."


script_end_procedure("")








