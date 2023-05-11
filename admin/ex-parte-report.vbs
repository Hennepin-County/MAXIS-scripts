'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EX PARTE REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("05/10/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================


'END FUNCTIONS BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================


'END DECLARATIONS BLOCK ====================================================================================================






'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Confirm_Process_to_Run_btn	= 200
incorrect_process_btn		= 100

If Day(date) < 1 Then ex_parte_function = "Prep"

'DISPLAYS DIALOG

DO
	DO
		DO
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 401, 255, "Ex Parte Report"
				DropListBox 300, 25, 90, 15, "Select one..."+chr(9)+"Prep"+chr(9)+"Phase 1"+chr(9)+"Phase 2", ex_parte_function
				ButtonGroup ButtonPressed
					OkButton 290, 235, 50, 15
					CancelButton 345, 235, 50, 15
				Text 5, 10, 400, 10, "This script will connect to the SQL Table to pull a list of cases to operate on based on the Ex Parte functionality selected."
				Text 200, 30, 95, 10, "Selection Ex Parte Function:"
				Text 10, 45, 35, 10, "Prep"
				Text 50, 45, 150, 10, "Timing - 4 Days before the BUDGET Month"
				Text 50, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 50, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 50, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
				Text 10, 90, 35, 10, "Phase 1"
				Text 50, 90, 135, 10, "Timing - 1st Day of the BUDGET Month"
				Text 50, 100, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 50, 110, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 50, 120, 125, 10, "Run each case through Background."
				Text 50, 130, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 50, 140, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
				Text 10, 155, 35, 10, "Phase 2"
				Text 50, 155, 160, 10, "Timing - 1st Day of the PROCESSING Month"
				Text 50, 165, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 50, 175, 145, 10, "Record in SQL Table any Updates found."
				Text 50, 185, 125, 10, "Run each case through Background."
				Text 50, 195, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 10, 215, 205, 10, "* * * * * THIS SCRIPT MUST BE RUN IN PRODUCTION * * * * *"
				Text 10, 235, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			EndDialog

			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
			If ex_parte_function = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an Ex Parte Function."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP until err_msg = ""

		If ex_parte_function = "Prep" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)


		End If
		If ex_parte_function = "Phase 1" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 2, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 2, date)), 2)

		End If
		If ex_parte_function = "Phase 2" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 1, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 1, date)), 2)

		End If

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 341, 165, "Confirm Ex Parte process"
			EditBox 600, 700, 10, 10, fake_edit_box
			ButtonGroup ButtonPressed
				PushButton 10, 145, 210, 15, "CONFIRMED! This is the correct Process and Review Month", Confirm_Process_to_Run_btn
				PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
			Text 10, 10, 225, 10, "You are running the Ex Parte Function " & ex_parte_function
			Text 10, 25, 190, 10, "This will run for the Ex Parte Review month of " & ep_revw_mo & "/" & ep_revw_yr
			If ex_parte_function = "Prep" Then
				GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
				Text 20, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 20, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 20, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
			End If
			If ex_parte_function = "Phase 1" Then
				GroupBox 5, 40, 295, 70, "Tasks to be Completed:"
				Text 20, 55, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 20, 65, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 20, 95, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
			End If
			If ex_parte_function = "Phase 2" Then
				GroupBox 5, 40, 305, 60, "Tasks to be Completed:"
				Text 20, 55, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 20, 65, 145, 10, "Record in SQL Table any Updates found."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
			End If
			Text 10, 115, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			Text 10, 130, 330, 10, "Review the process datails and ex parte review month to confirm this is the correct run to complete."
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed = OK Then ButtonPressed = Confirm_Process_to_Run_btn

	Loop until ButtonPressed = Confirm_Process_to_Run_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in





Cal script_end_procedure("DONE")