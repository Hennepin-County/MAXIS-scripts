'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
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

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result)
	Call write_value_and_transmit("99", cmd_row, cmd_col)

	row = 7
	Do
		EMReadScreen elig_version, 2, row, 22
		EmReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do

		row = row + 1
	Loop until approval_status = ""
	Call clear_line_of_text(18, 54)

	Call write_value_and_transmit(elig_version, 18, 54)
	version_number = "0" & elig_version
	version_date = elig_date
	version_result = elig_result
end if

function read_SNAP_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "FS  ")
	Call find_last_approved_ELIG_version(19, 78, version_number, version_date, elig_result)

	row = 7
	Do
		EMReadScreen ref_numb, 2, row, 6
		If ref_numb <> "  " Then
			EMReadScreen request_yn, 1, row, 32
			EMReadScreen memb_code, 1, row, 36
			EMReadScreen memb_count, 11, row, 41
			EMReadScreen memb_elig, 10, row, 53
			EMReadScreen memb_begin_date, 8, row, 67
			EMReadScreen memb_budg_cycle, 1, row, 78

			Call write_value_and_transmit("X", row, 3)
			EMReadScreen memb_absence, 			6, 7, 17
			EMReadScreen memb_child_age, 		6, 8, 17
			EMReadScreen memb_citizenship, 		6, 9, 17
			EMReadScreen memb_citizenship_ver, 	6, 10, 17
			EMReadScreen memb_dup_assist, 		6, 11, 17
			EMReadScreen memb_fost_care, 		6, 12, 17
			EMReadScreen memb_fraud, 			6, 13, 17
			EMReadScreen memb_disq, 			6, 17, 17

			EMReadScreen memb_minor_living, 6, 7, 52
			EMReadScreen memb_post_60, 6, 7, 52
			EMReadScreen memb_ssi, 6, 7, 52
			EMReadScreen memb_ssn_coop, 6, 7, 52
			EMReadScreen memb_unit_memb, 6, 7, 52
			EMReadScreen memb_unlawful_conduct, 6, 7, 52
			EMReadScreen memb_fs_recvd, 6, 7, 52

		End If

		row = row + 1
	Loop until ref_numb = "  "

	Call Back_to_SELF
end Function


function read_MFIP_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "MFIP")
	Call find_last_approved_ELIG_version(20, 79, version_number, version_date, elig_result)


	Call Back_to_SELF
end Function


function read_DWP_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "DWP ")



	Call Back_to_SELF
end Function


function read_GA_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "GA  ")



	Call Back_to_SELF
end Function

function read_MSA_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "MSA ")



	Call Back_to_SELF
end Function

function read_GRH_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "GRH ")



	Call Back_to_SELF
end Function

function read_MA_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "HC  ")



	Call Back_to_SELF
end Function

function read_MSP_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "HC  ")



	Call Back_to_SELF
end Function

function read_EMER_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "EMER")



	Call Back_to_SELF
end Function

function read_CASH_elig(elig_footer_month, elig_footer_year)
	call navigate_to_MAXIS_screen("ELIG", "DENY")



	Call Back_to_SELF
end Function



















































'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
