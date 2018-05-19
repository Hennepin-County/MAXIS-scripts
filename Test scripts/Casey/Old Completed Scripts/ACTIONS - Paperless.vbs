'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - paperless"
start_time = timer
STATS_counter = 0              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.  
STATS_manualtime = 60          'manual run time in seconds
STATS_denomination = "C"       'I is for each dail message 

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


'DATE/TIME CALCULATIONS

current_day = DatePart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day 
current_month = DatePart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month 
current_year = DatePart("yyyy", date)
current_year = current_year - 2000

'THE SCRIPT


EMConnect ""

continue_prompt = MsgBox("***AT THIS TIME, THIS SCRIPT IS ONLY FOR ADULT WORKERS***"& Chr(13) & Chr(13) &_
"This script will update REVW for each single-adult starred IR, after checking JOBS/BUSI/RBIC for discrepancies. It skips cases that are also reviewing for SNAP." & Chr(13) &_
"You will have to manually check elig/HC for each case and approve the results/case note. Press OK to begin!", 1, "Are you sure?")
If continue_prompt = 2 then stopscript

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
  if paperless_check = "*" then MAXIS_case_number_array = trim(MAXIS_case_number_array & " " & trim(MAXIS_case_number))
  row = row + 1
Loop until last_page_check = "LAST" or trim(MAXIS_case_number) = ""


MAXIS_case_number_array = split(MAXIS_case_number_array)

For each MAXIS_case_number in MAXIS_case_number_array
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
  
  call navigate_to_MAXIS_screen ("stat", "busi") 'Updated 08/26/15 - to match FuncLib - CLove
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

  call navigate_to_MAXIS_screen ("stat", "rbic") 'Updated 08/26/15 - to match FuncLib - CLove
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
    call navigate_to_MAXIS_screen ("stat", "revw") 'Updated 08/26/15 - to match FuncLib - CLove
    EMReadScreen SNAP_review_check, 1, 7, 60
    If SNAP_review_check <> "N" then
	  STATS_counter = STATS_counter + 1
	  cases_to_tikl = cases_to_tikl & "~" & MAXIS_case_number
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
      EMWriteScreen current_month, 6, 27
      EMWriteScreen current_day, 6, 30
      EMWriteScreen current_year, 6, 33
      new_renewal_year = cint(current_year) + 1
      If current_month = 12 then new_renewal_year = new_renewal_year + 1 'Becuase otherwise the renewal year will be the current footer month.
      EMWriteScreen new_renewal_year, 8, renewal_year_col
      EMWriteScreen "U", 13, 43
      EMReadScreen spouse_check, 1, 14, 43
      If spouse_check = "N" then PF10
      transmit
    End if
  End if
Next

If cases_to_tikl <> "" Then 
	cases_to_tikl = right(cases_to_tikl, len(cases_to_tikl)-1)
	cases_to_tikl_array = split(cases_to_tikl, "~")
End If 

For each MAXIS_case_number in cases_to_tikl_array
	navigate_to_MAXIS_screen "DAIL", "WRIT"
	EMWritescreen current_month, 5, 18
	EMWritescreen current_day,   5, 21
	EMWritescreen current_year,  5, 24
	transmit
	EMWritescreen "%^% Sent through background using bulk script %^%", 9, 3
	transmit
	EMReadScreen tikl_success, 4, 24, 2
	If tikl_success <> "    " Then MsgBox "This case - " & MAXIS_case_number & " failed to have a TIKL set, track and case note manually"
	PF3
Next

transmit

Do
  PF3
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

script_end_procedure("Success! All starred (*) IRs have been sent into background, except those with current JOBS/BUSI/RBIC, those who have members other than 01 open, or those who also have SNAP up for review. " & chr(13) + _
"You must go through and approve these results when they come through background. Talk to a PC or supervisor if you have any questions about paperless policy.") 'Moved MsgBox to Script End Proceedure - CLove