'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - PAPERLESS Review.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1            'sets the stats counter at one
STATS_manualtime = 0         'manual run time in seconds
STATS_denomination = "C"     'C is for each case
'END OF stats block=========================================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog paperless_IR_dialog, 0, 0, 136, 70, "NON-MAGI PAPERLESS IR"
  EditBox 75, 5, 55, 15, worker_number
  EditBox 75, 25, 25, 15, MAXIS_footer_month
  EditBox 105, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 50, 50, 15
    CancelButton 80, 50, 50, 15
  Text 5, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 55, 10, "Worker number:"
EndDialog

'establishing variable for the script since most users are approving CM + 1
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
'other dates for script

current_month = CM_mo
current_day = "01" 
current_year = CM_yr
'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'msgbox current_month & current_day & current_year
'Script warning, this is only for adult workers (non- magi) at this time				
continue_prompt = MsgBox("***AT THIS TIME, THIS SCRIPT IS ONLY FOR ADULT/ADS WORKERS***"& Chr(13) & Chr(13) &_
"This script will update REVW for each adult starred IR, after checking JOBS/BUSI/RBIC for discrepancies. It skips cases that are also reviewing for SNAP." & Chr(13) &_
"You will have to manually check elig/HC for each case and approve the results/case note. Press OK to begin!", 1, "Are you sure?")
If continue_prompt = 2 then stopscript

DO
	DO	
		err_msg = ""
		Dialog paperless_IR_dialog
		If buttonpressed = 0 then stopscript
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If Len(worker_number) <> 7 then err_msg = err_msg & vbNewLine & "* You must enter a valid 7 DIGIT worker number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS				
Loop until are_we_passworded_out = false					'loops until user passwords back in			

Call navigate_to_MAXIS_screen("rept", "revw")
Call MAXIS_footer_month_confirmation
EMWriteScreen worker_number, 21, 6
transmit

EMReadScreen REVW_check, 4, 2, 52
If REVW_check <> "REVW" then script_end_procedure("You must start this script at the beginning of REPT/REVW. Navigate to the screen and try again!")

row = 7
Do
  If row = 19 then
    PF8
    row = 7
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then stopscript
    EMReadScreen last_page_check, 4, 24, 14
  End if
  EMReadScreen MAXIS_case_number, 8, row, 6
  EMReadScreen paperless_check, 1, row, 51
  if paperless_check = "*" then case_number_array = trim(case_number_array & " " & trim(MAXIS_case_number))
  row = row + 1
Loop until last_page_check = "LAST" or trim(MAXIS_case_number) = ""

case_number_array = split(case_number_array)

For each MAXIS_case_number in case_number_array
  actually_paperless = "" 'Resetting the variable.
  call navigate_to_MAXIS_screen ("stat", "memb")
  call navigate_to_MAXIS_screen ("stat", "jobs")
  EMWriteScreen "01", 20, 76
  transmit
  Do
    EMReadScreen panel_check, 8, 2, 72
    current_panel = trim(left(panel_check, 2))
    total_panels = trim(right(panel_check, 2))
    EMReadScreen date_check, 8, 9, 49
    If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
    if current_panel <> total_panels then transmit
  Loop until current_panel = total_panels
  
  call navigate_to_MAXIS_screen ("stat", "busi")
  EMWriteScreen "01", 20, 76
  transmit
  Do
    EMReadScreen panel_check, 8, 2, 72
    current_panel = trim(left(panel_check, 2))
    total_panels = trim(right(panel_check, 2))
    EMReadScreen date_check, 8, 5, 71
    If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
    if current_panel <> total_panels then transmit
  Loop until current_panel = total_panels

  call navigate_to_MAXIS_screen ("stat", "rbic")
  EMWriteScreen "01", 20, 76
  transmit
  Do
    EMReadScreen panel_check, 8, 2, 72
    current_panel = trim(left(panel_check, 2))
    total_panels = trim(right(panel_check, 2))
    EMReadScreen date_check, 8, 6, 68
    If total_panels <> "0" & date_check = "__ __ __" then actually_paperless = False
    if current_panel <> total_panels then transmit
  Loop until current_panel = total_panels

  If actually_paperless <> False then
    actually_paperless = True
  Else
    MsgBox "This case is not paperless!"
  End if

  If actually_paperless = True then
    call navigate_to_MAXIS_screen ("stat", "revw")
    EMReadScreen SNAP_review_check, 1, 7, 60
    If SNAP_review_check <> "N" then
      PF9
      EMWriteScreen "x", 5, 71
      transmit
      EMReadScreen renewal_year, 2, 8, 33
      If renewal_year = "__" then
        EMReadScreen renewal_year, 2, 8, 77
        renewal_year_col = 77
      Else
        renewal_year_col = 33
      End if
      EMWriteScreen left(current_month, 2), 6, 27
      EMWriteScreen current_day, 6, 30
      EMWriteScreen right(current_year, 2), 6, 33
      new_renewal_year = cint(right(current_year, 2)) + 1
      If current_month = 12 then new_renewal_year = new_renewal_year + 1 'Because otherwise the renewal year will be the current footer month.
      EMWriteScreen new_renewal_year, 8, renewal_year_col
      EMWriteScreen "U", 13, 43
      EMReadScreen spouse_check, 1, 14, 43
      If spouse_check = "N" then PF10
      transmit
    End if
  End if
Next

transmit

Do
  PF3
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

script_end_procedure("Success! All starred (*) IRs have been sent into background, except those with current JOBS/BUSI/RBIC, those who have members other than 01 open, or those who also have SNAP up for review. You must go through and approve these results when they come through background.")